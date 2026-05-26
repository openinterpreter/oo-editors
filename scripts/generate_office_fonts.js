#!/usr/bin/env node

/**
 * Refresh ONLYOFFICE font metadata in assets/onlyoffice-fontdata.
 *
 * Requires:
 *   - converter/allfontsgen (compiled previously)
 *   - converter/{graphics,kernel,UnicodeConverter}.framework
 */

const { spawnSync } = require('child_process');
const fs = require('fs');
const path = require('path');

function fail(msg) {
  console.error(`[generate_office_fonts] ${msg}`);
  process.exit(1);
}

const ROOT = path.resolve(__dirname, '..');
const CONVERTER_DIR = path.join(ROOT, 'converter');
const platform = process.platform;

if (!process.env.FONT_DATA_DIR) {
  fail('FONT_DATA_DIR environment variable is required');
}

const OUTPUT_DIR = process.env.FONT_DATA_DIR;

// Early return if font metadata files already exist
function checkExistingFonts() {
  const allFontsJs = path.join(OUTPUT_DIR, 'AllFonts.js');
  const fontSelectionBin = path.join(OUTPUT_DIR, 'font_selection.bin');

  if (fs.existsSync(allFontsJs) && fs.existsSync(fontSelectionBin)) {
    console.log('[generate_office_fonts] Font metadata already exists in', OUTPUT_DIR);
    return true;
  }
  return false;
}

if (checkExistingFonts()) {
  process.exit(0);
}

function locateBinary() {
  if (platform === 'darwin') {
    const macBin = path.join(CONVERTER_DIR, 'allfontsgen');
    if (fs.existsSync(macBin)) return macBin;
    fail(`macOS binary missing: ${macBin}. Run scripts/build_allfontsgen.js first.`);
  } else if (platform === 'win32') {
    const winBin = path.join(CONVERTER_DIR, 'allfontsgen.exe');
    if (fs.existsSync(winBin)) return winBin;
    fail(`Windows binary missing: ${winBin}. Build it and place it under converter/.`);
  } else if (platform === 'linux') {
    const linuxBin = path.join(CONVERTER_DIR, 'allfontsgen');
    if (fs.existsSync(linuxBin)) return linuxBin;
    fail(`Linux binary missing: ${linuxBin}. Build it and place it under converter/.`);
  }
  fail(`Unsupported platform: ${platform}`);
}

const BIN = locateBinary();

const INPUT_DIR = path.join(OUTPUT_DIR, 'fonts');
const DIAG_LOG = path.join(OUTPUT_DIR, 'allfontsgen-control.log');

fs.mkdirSync(OUTPUT_DIR, { recursive: true });
fs.mkdirSync(INPUT_DIR, { recursive: true });

const args = [
  '--use-system=true',
  `--input=${INPUT_DIR}`,
  `--allfonts=${path.join(OUTPUT_DIR, 'AllFonts.js')}`,
  `--selection=${path.join(OUTPUT_DIR, 'font_selection.bin')}`,
];

const env = { ...process.env };
if (platform === 'darwin') {
  env.DYLD_FRAMEWORK_PATH = CONVERTER_DIR;
}
if (platform === 'win32') {
  env.ALLFONTSGEN_DIAG_LOG = DIAG_LOG;
}

function readDiagnosticLog() {
  if (platform !== 'win32' || !fs.existsSync(DIAG_LOG)) {
    return '';
  }

  const text = fs.readFileSync(DIAG_LOG, 'utf8').trim();
  if (!text) {
    return '';
  }

  const lines = text.split(/\r?\n/).slice(-20);
  return `\n[generate_office_fonts] allfontsgen diagnostic log:\n${lines.join('\n')}`;
}

function formatExitStatus(status) {
  if (typeof status !== 'number') {
    return String(status);
  }

  return `${status} (0x${(status >>> 0).toString(16).toUpperCase()})`;
}

const result = spawnSync(BIN, args, { stdio: 'inherit', env });

if (result.status !== 0) {
  const details = [
    `allfontsgen exited with code ${formatExitStatus(result.status)}`,
    result.signal ? `signal=${result.signal}` : null,
    result.error ? `error=${result.error.message}` : null,
  ].filter(Boolean).join(' ');

  fail(`${details}${readDiagnosticLog()}`);
}

const allFontsJs = path.join(OUTPUT_DIR, 'AllFonts.js');
const fontSelectionBin = path.join(OUTPUT_DIR, 'font_selection.bin');

if (!fs.existsSync(allFontsJs) || !fs.existsSync(fontSelectionBin)) {
  const missing = [
    fs.existsSync(allFontsJs) ? null : allFontsJs,
    fs.existsSync(fontSelectionBin) ? null : fontSelectionBin,
  ].filter(Boolean);

  fail(
    `allfontsgen exited successfully but font metadata files were not created; missing=${missing.join(', ')}${readDiagnosticLog()}`
  );
}

console.log('[generate_office_fonts] Font metadata updated.');
process.exit(0);
