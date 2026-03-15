#!/usr/bin/env node

/**
 * generate-manifests.js
 *
 * Reads manifest templates from manifests-templates/, replaces
 * {{FRONTEND_URL}} and {{BACKEND_URL}} placeholders with values
 * built from environment variables, and writes the resulting
 * manifest files to generated-manifests/ as pure UTF-8 (no BOM).
 *
 * Environment variables:
 *   SERVER_IP       – host IP (required)
 *   FRONTEND_PORT   – frontend port (default: 3002)
 *   BACKEND_PORT    – backend port  (default: 3003)
 *
 * -------------------------------------------------------------------
 * ARCH-L2: Manifest Serving Strategy
 * -------------------------------------------------------------------
 *
 * CURRENT APPROACH (Self-hosted / Internal):
 *   - Manifests generated to: generated-manifests/ (project root)
 *   - Served via Express route: /manifests/manifest-office.xml
 *   - Benefits: Can add authentication, rate limiting, server-side logic
 *   - Security: Manifests contain internal hostnames/URLs but are only
 *               accessible to authenticated users on the internal network
 *
 * ALTERNATIVE APPROACH (SaaS / Public distribution):
 *   - Output manifests to: frontend/public/assets/manifests/
 *   - Served as static files directly by Vite/Nginx
 *   - Benefits: Works with static hosting (CDN), no Express dependency,
 *               same-origin serving (no CORS), simpler distribution
 *   - Security considerations:
 *       * Manifests become publicly discoverable
 *       * Internal hostnames/URLs exposed to anyone with the URL
 *       * Mitigation: Use relative paths where possible, serve manifests
 *                     only at non-obvious paths, implement allowlist for
 *                     which configurations can be served publicly
 *
 * RECOMMENDATION:
 *   - Keep current approach for self-hosted/internal deployments
 *   - When moving to SaaS model, change OUTPUT_DIR to:
 *     path.join(ROOT_DIR, 'frontend/public/assets/manifests')
 *   - Add route-level authentication/allowlist before exposing publicly
 * -------------------------------------------------------------------
 */

'use strict';

const fs   = require('fs');
const path = require('path');

// ---------------------------------------------------------------------------
// 1. Resolve environment variables
// ---------------------------------------------------------------------------

const SERVER_IP     = process.env.SERVER_IP;
const FRONTEND_PORT = process.env.FRONTEND_PORT || '3002';
const BACKEND_PORT  = process.env.BACKEND_PORT  || '3003';

if (!SERVER_IP) {
  console.error('ERROR: SERVER_IP environment variable is not set.');
  process.exit(1);
}

const FRONTEND_URL = process.env.PUBLIC_FRONTEND_URL || `http://${SERVER_IP}:${FRONTEND_PORT}`;
const BACKEND_URL  = process.env.PUBLIC_BACKEND_URL  || `http://${SERVER_IP}:${BACKEND_PORT}`;

// ---------------------------------------------------------------------------
// 2. Paths
// ---------------------------------------------------------------------------

const ROOT_DIR      = path.resolve(__dirname, '..');
const TEMPLATES_DIR = path.join(ROOT_DIR, 'manifests-templates');
const OUTPUT_DIR    = path.join(ROOT_DIR, 'generated-manifests');

fs.mkdirSync(OUTPUT_DIR, { recursive: true });

// ---------------------------------------------------------------------------
// 2b. Ensure backend log directories exist with correct ownership
//
// The backend bind-mount (./backend/logs → /app/logs) requires the host
// directories to exist BEFORE the backend container starts, and to be
// writable by the node user (UID/GID 1000) inside the container.
// Creating them here (this script runs as root in manifest-gen, which
// already runs before kickoffice-backend) avoids a dedicated logs-init
// container.
//
// IMPORTANT: manifest-gen must mount ./backend/logs:/app/logs for this to work.
// ---------------------------------------------------------------------------

const LOGS_DIR = '/app/logs';
const LOG_SUBDIRS = ['feedback', 'frontend'];
for (const sub of LOG_SUBDIRS) {
  fs.mkdirSync(path.join(LOGS_DIR, sub), { recursive: true });
}

// Set ownership to UID/GID 1000 (node user inside the container).
// This mirrors what the former logs-init container did.  Only attempt
// chown when running as root (CI/CD, Docker) — skip silently otherwise.
try {
  const { execSync } = require('child_process');
  execSync(`chown -R 1000:1000 ${LOGS_DIR}`);
  console.log('[manifest-gen] Log directories created and ownership set to 1000:1000');
} catch {
  console.log('[manifest-gen] Log directories created (chown skipped — not running as root)');
}

const templates = [
  {
    src:  path.join(TEMPLATES_DIR, 'manifest-office.template.xml'),
    dest: path.join(OUTPUT_DIR, 'manifest-office.xml'),
    label: 'Office (Word / Excel / PowerPoint)',
  },
  {
    src:  path.join(TEMPLATES_DIR, 'manifest-outlook.template.xml'),
    dest: path.join(OUTPUT_DIR, 'manifest-outlook.xml'),
    label: 'Outlook',
  },
];

// ---------------------------------------------------------------------------
// 3. Generate
// ---------------------------------------------------------------------------

for (const { src, dest, label } of templates) {
  if (!fs.existsSync(src)) {
    console.error(`ERROR: Template not found: ${src}`);
    process.exit(1);
  }

  let content = fs.readFileSync(src, 'utf8');

  // Strip BOM if present in the template
  if (content.charCodeAt(0) === 0xFEFF) {
    content = content.slice(1);
  }

  content = content
    .replace(/\{\{FRONTEND_URL\}\}/g, FRONTEND_URL)
    .replace(/\{\{BACKEND_URL\}\}/g, BACKEND_URL);

  fs.writeFileSync(dest, content, { encoding: 'utf8' });
  console.log(`[manifest-gen] ${label} -> ${path.basename(dest)}  (${FRONTEND_URL})`);
}

console.log(`[manifest-gen] Done – generated ${templates.length} manifest(s) for ${SERVER_IP}`);
