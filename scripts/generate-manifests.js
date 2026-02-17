#!/usr/bin/env node

/**
 * generate-manifests.js
 *
 * Reads manifest templates from manifests-templates/, replaces
 * {{FRONTEND_URL}} / {{SERVER_URL}} placeholders with values
 * built from environment variables, and writes the resulting
 * manifest files to the project root as pure UTF-8 (no BOM).
 *
 * Environment variables:
 *   FRONTEND_URL    – public frontend URL (preferred)
 *
 * Legacy fallback environment variables (for local/dev compatibility):
 *   SERVER_IP       – host IP (required when FRONTEND_URL is not set)
 *   FRONTEND_PORT   – frontend port (default: 3002)
 */

'use strict';

const fs   = require('fs');
const path = require('path');

function loadRootDotEnv() {
  const rootEnvPath = path.resolve(__dirname, '..', '.env');
  if (!fs.existsSync(rootEnvPath)) {
    return;
  }

  const envContent = fs.readFileSync(rootEnvPath, 'utf8');
  for (const line of envContent.split(/\r?\n/)) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith('#')) {
      continue;
    }

    const separatorIndex = trimmed.indexOf('=');
    if (separatorIndex <= 0) {
      continue;
    }

    const key = trimmed.slice(0, separatorIndex).trim();
    const rawValue = trimmed.slice(separatorIndex + 1).trim();
    const value = rawValue.replace(/^['"]|['"]$/g, '');

    if (!process.env[key]) {
      process.env[key] = value;
    }
  }
}

loadRootDotEnv();

// ---------------------------------------------------------------------------
// 1. Resolve environment variables
// ---------------------------------------------------------------------------

const SERVER_IP = process.env.SERVER_IP;
const FRONTEND_PORT = process.env.FRONTEND_PORT || '3002';
const FRONTEND_URL = (process.env.FRONTEND_URL || '').trim() || (SERVER_IP ? `http://${SERVER_IP}:${FRONTEND_PORT}` : '');

if (!FRONTEND_URL) {
  console.error('ERROR: FRONTEND_URL is required (or provide SERVER_IP as legacy fallback).');
  process.exit(1);
}

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
    .replace(/\{\{SERVER_URL\}\}/g, FRONTEND_URL)
    .replace(/\{\{BACKEND_URL\}\}/g, FRONTEND_URL);

  fs.writeFileSync(dest, content, { encoding: 'utf8' });
  console.log(`[manifest-gen] ${label} -> ${path.basename(dest)}  (${FRONTEND_URL})`);
}

console.log(`[manifest-gen] Done – generated ${templates.length} manifest(s) for ${FRONTEND_URL}`);
