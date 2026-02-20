#!/usr/bin/env node

/**
 * bump-version.js
 *
 * Automatically increments the patch version of the KickOffice application.
 * It reads `frontend/package.json`, bumps the patch number (e.g. 1.0.0 -> 1.0.1),
 * applies it to `backend/package.json`, and updates the Office manifestations
 * (`manifests-templates/*.xml` and root `manifest-*.xml`).
 *
 * Usage: node scripts/bump-version.js
 */

const fs = require('fs');
const path = require('path');

const ROOT_DIR = path.resolve(__dirname, '..');
const FRONTEND_PKG_PATH = path.join(ROOT_DIR, 'frontend', 'package.json');
const BACKEND_PKG_PATH = path.join(ROOT_DIR, 'backend', 'package.json');
const MANIFESTS_TEMPLATES_DIR = path.join(ROOT_DIR, 'manifests-templates');

const TEMPLATES = [
  path.join(MANIFESTS_TEMPLATES_DIR, 'manifest-office.template.xml'),
  path.join(MANIFESTS_TEMPLATES_DIR, 'manifest-outlook.template.xml'),
  path.join(ROOT_DIR, 'manifest-office.xml'),
  path.join(ROOT_DIR, 'manifest-outlook.xml')
];

function bumpVersion(versionStr) {
  const parts = versionStr.split('.');
  if (parts.length !== 3) {
    throw new Error(`Invalid version format in package.json: ${versionStr}`);
  }
  const major = parseInt(parts[0], 10);
  const minor = parseInt(parts[1], 10);
  const patch = parseInt(parts[2], 10);
  
  return `${major}.${minor}.${patch + 1}`;
}

function updateJson(filePath, newVersion) {
  if (!fs.existsSync(filePath)) {
    console.warn(`[WARNING] File not found: ${filePath}`);
    return;
  }
  const content = JSON.parse(fs.readFileSync(filePath, 'utf8'));
  content.version = newVersion;
  fs.writeFileSync(filePath, JSON.stringify(content, null, 2) + '\n', 'utf8');
  console.log(`[Success] Updated ${path.basename(path.dirname(filePath))}/${path.basename(filePath)} to v${newVersion}`);
}

function updateXml(filePath, newVersion) {
  if (!fs.existsSync(filePath)) {
    console.warn(`[WARNING] File not found: ${filePath}`);
    return;
  }
  
  // Office manifests require a 4-part version (e.g., 1.0.1.0)
  const officeVersion = `${newVersion}.0`;
  
  let content = fs.readFileSync(filePath, 'utf8');
  
  // Strip BOM if present
  if (content.charCodeAt(0) === 0xFEFF) {
    content = content.slice(1);
  }

  // Regex to match <Version>X.X.X.X</Version>
  const versionRegex = /<Version>[\d\.]+<\/Version>/i;
  
  if (!versionRegex.test(content)) {
    console.warn(`[WARNING] <Version> tag not found in ${filePath}`);
  }

  content = content.replace(versionRegex, `<Version>${officeVersion}</Version>`);
  
  fs.writeFileSync(filePath, content, { encoding: 'utf8' });
  console.log(`[Success] Updated ${path.basename(filePath)} <Version> to ${officeVersion}`);
}

function run() {
  console.log('--- KickOffice Version Bumper ---');

  if (!fs.existsSync(FRONTEND_PKG_PATH)) {
    console.error(`[ERROR] Primary source of truth not found: ${FRONTEND_PKG_PATH}`);
    process.exit(1);
  }

  const frontendPkg = JSON.parse(fs.readFileSync(FRONTEND_PKG_PATH, 'utf8'));
  const currentVersion = frontendPkg.version || '1.0.0';
  
  console.log(`Current version: v${currentVersion}`);
  
  let newVersion;
  try {
    newVersion = bumpVersion(currentVersion);
  } catch (err) {
    console.error(`[ERROR] ${err.message}`);
    process.exit(1);
  }

  console.log(`Target version:  v${newVersion}`);
  console.log('---------------------------------');

  updateJson(FRONTEND_PKG_PATH, newVersion);
  updateJson(BACKEND_PKG_PATH, newVersion);

  for (const xmlPath of TEMPLATES) {
    updateXml(xmlPath, newVersion);
  }

  console.log('---------------------------------');
  console.log(`✅ Bumped all files successfully to v${newVersion}`);
}

run();
