#!/usr/bin/env node

/**
 * generate-manifests.js
 *
 * Reads manifest templates from manifests-templates/, replaces
 * {{FRONTEND_URL}} and {{BACKEND_URL}} placeholders with values
 * built from environment variables, and writes the resulting
 * manifest files to the project root as pure UTF-8 (no BOM).
 *
 * Environment variables:
 *   SERVER_IP       – host IP (required)
 *   FRONTEND_PORT   – frontend port (default: 3002)
 *   BACKEND_PORT    – backend port  (default: 3003)
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

const FRONTEND_URL = `http://${SERVER_IP}:${FRONTEND_PORT}`;
const BACKEND_URL  = `http://${SERVER_IP}:${BACKEND_PORT}`;

// ---------------------------------------------------------------------------
// 2. Paths
// ---------------------------------------------------------------------------

const ROOT_DIR      = path.resolve(__dirname, '..');
const TEMPLATES_DIR = path.join(ROOT_DIR, 'manifests-templates');

const templates = [
  {
    src:  path.join(TEMPLATES_DIR, 'manifest-office.template.xml'),
    dest: path.join(ROOT_DIR, 'manifest-office.xml'),
    label: 'Office (Word / Excel / PowerPoint)',
  },
  {
    src:  path.join(TEMPLATES_DIR, 'manifest-outlook.template.xml'),
    dest: path.join(ROOT_DIR, 'manifest-outlook.xml'),
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
