#!/usr/bin/env bun

const https = require('https');
const fs = require('fs');
const path = require('path');
const { execSync, spawnSync } = require('child_process');
const os = require('os');

// NOTE(victor): Windows MUST use x64 binaries. x86 fails with 0xC0000135 (missing VC++ runtime).
// Available builds: https://github.com/ONLYOFFICE/DesktopEditors/releases
//
// IMPORTANT: When changing this version, you MUST also update the sdkjs/ subtree and
// editors/web-apps/ to match. The DesktopEditors release pins specific commits of
// ONLYOFFICE/sdkjs and ONLYOFFICE/web-apps. Mismatched versions cause subtle bugs
// (wrong font rendering, broken chart serialization, API mismatches between SDK and UI).
//
// To find the sdkjs commit for a given release:
//   gh api repos/ONLYOFFICE/DesktopEditors/git/trees/<tag> \
//     --jq '.tree[] | select(.path=="sdkjs") | .sha'
//
// Then re-pin the subtree:
//   git rm -r sdkjs/ && rm -rf sdkjs/
//   git commit -m "chore: remove sdkjs subtree for re-pin"
//   git subtree add --prefix=sdkjs https://github.com/ONLYOFFICE/sdkjs.git <commit> --squash
//
// v9.1.0 pins: sdkjs @ d169f841a7e9e46368c36236dd5820e3e10d4a98
//
// Project files modified from upstream (excluding test/CI/docs):
//
//   Server / Core:
//     server.js                    - Express server (dynamic port, default 38123)
//     server-utils.js              - Server utility functions (added)
//     download-converter.js        - Downloads ONLYOFFICE Desktop, extracts x2t (added)
//     package.json                 - Dependencies and scripts
//     bun.lock                     - Lockfile (added)
//     .gitignore                   - Ignore patterns
//
//   Editors:
//     editors/desktop-stub.js                              - Browser stub for AscDesktopEditor API
//     editors/desktop-stub-utils.js                        - Utility helpers for desktop stub (added)
//     editors/offline-loader-proper.html                   - Document loading orchestration
//     editors/plugins.json                                 - Plugin configuration (added)
//     editors/sdkjs/common/Local/common.js                 - Patched SDK local mode for browser/HTTP
//     editors/web-apps/apps/api/documents/cache-scripts.html - Script caching
//     editors/web-apps/apps/api/documents/preload.html     - Preload configuration
//     editors/web-apps/vendor/socketio/socket.io.min.js    - Socket.io vendor update
//
//   Scripts:
//     scripts/build_allfontsgen.js      - Builds font metadata generator (added)
//     scripts/generate_office_fonts.js  - Creates AllFonts.js and font_selection.bin (added)
//
//   sdkjs/ patches (NOTE(victor) markers):
//     sdkjs/common/Local/common.js:185  - URL protocol check fix (file: vs http://)
//     sdkjs/common/libfont/map.js:2933  - Guard for documents referencing missing fonts
//     sdkjs/common/libfont/map.js:2969  - GetFontFileWeb font name resolution
const BASE_URL = 'https://github.com/ONLYOFFICE/DesktopEditors/releases/download/v9.1.0';
const WINDOWS_ZIP_NAME = 'DesktopEditors_x64.zip';
const CONVERTER_DIR = path.join(__dirname, 'converter');

function getDownloadUrl() {
  const platform = process.env.TARGET_PLATFORM || os.platform();
  const arch = process.env.TARGET_ARCH || os.arch();

  if (platform === 'darwin') {
    return arch === 'arm64'
      ? `${BASE_URL}/ONLYOFFICE-arm.dmg`
      : `${BASE_URL}/ONLYOFFICE-x86_64.dmg`;
  }

  if (platform === 'win32') {
    return `${BASE_URL}/${WINDOWS_ZIP_NAME}`;
  }

  throw new Error(`Unsupported platform: ${platform}`);
}

function download(url, dest) {
  return new Promise((resolve, reject) => {
    const file = fs.createWriteStream(dest);

    https.get(url, (response) => {
      if (response.statusCode === 302 || response.statusCode === 301) {
        file.close();
        fs.unlinkSync(dest);
        return download(response.headers.location, dest).then(resolve).catch(reject);
      }

      if (response.statusCode !== 200) {
        file.close();
        fs.unlinkSync(dest);
        return reject(new Error(`Download failed: ${response.statusCode}`));
      }

      response.pipe(file);

      file.on('finish', () => {
        file.close();
        resolve();
      });
    }).on('error', (err) => {
      fs.unlinkSync(dest);
      reject(err);
    });
  });
}

function extractDmg(dmgPath) {
  const mountOutput = execSync(`hdiutil attach -noverify -noautoopen "${dmgPath}"`, { encoding: 'utf8' });
  const volumePath = mountOutput.match(/\/Volumes\/[^\s]+/)?.[0];

  if (!volumePath) {
    throw new Error('Failed to mount DMG');
  }

  try {
    const converterSource = path.join(volumePath, 'ONLYOFFICE.app/Contents/Resources/converter');

    if (!fs.existsSync(converterSource)) {
      throw new Error(`Converter directory not found at ${converterSource}`);
    }

    fs.mkdirSync(CONVERTER_DIR, { recursive: true });
    execSync(`cp -R "${converterSource}"/* "${CONVERTER_DIR}"/`);
  } finally {
    execSync(`hdiutil detach "${volumePath}" -quiet`);
  }
}

function runPowerShellCommand(script) {
  const result = spawnSync(
    'powershell.exe',
    ['-NoProfile', '-NonInteractive', '-Command', script],
    { stdio: 'inherit' }
  );

  if (result.status !== 0) {
    throw new Error(`PowerShell command failed with exit code ${result.status || 1}`);
  }
}

function extractZip(zipPath) {
  const tempDir = path.join(__dirname, 'temp_extract');
  fs.rmSync(tempDir, { recursive: true, force: true });
  fs.mkdirSync(tempDir, { recursive: true });

  try {
    const psZip = zipPath.replace(/'/g, "''");
    const psTemp = tempDir.replace(/'/g, "''");
    runPowerShellCommand(
      `Expand-Archive -LiteralPath '${psZip}' -DestinationPath '${psTemp}' -Force`
    );

    const converterSource = findConverterDirectory(tempDir);

    if (!converterSource) {
      throw new Error('Converter directory not found in ZIP extraction');
    }

    fs.mkdirSync(CONVERTER_DIR, { recursive: true });
    execSync(`xcopy /E /I /Y "${converterSource}\\*" "${CONVERTER_DIR}"`, { stdio: 'inherit' });
  } finally {
    fs.rmSync(tempDir, { recursive: true, force: true });
  }
}

function findConverterDirectory(rootDir) {
  const queue = [rootDir];

  while (queue.length > 0) {
    const current = queue.pop();
    let entries = [];

    try {
      entries = fs.readdirSync(current, { withFileTypes: true });
    } catch {
      continue;
    }

    for (const entry of entries) {
      if (!entry.isDirectory()) continue;
      const fullPath = path.join(current, entry.name);
      if (entry.name.toLowerCase() === 'converter') {
        const kernelDll = path.join(fullPath, 'kernel.dll');
        if (fs.existsSync(kernelDll)) {
          return fullPath;
        }
      }
      queue.push(fullPath);
    }
  }

  return null;
}

function cleanup() {
  const emptyDir = path.join(CONVERTER_DIR, 'empty');
  const templatesDir = path.join(CONVERTER_DIR, 'templates');

  if (fs.existsSync(emptyDir)) {
    fs.rmSync(emptyDir, { recursive: true, force: true });
  }

  if (fs.existsSync(templatesDir)) {
    fs.rmSync(templatesDir, { recursive: true, force: true });
  }
}

async function main() {
  const x2tPath = path.join(CONVERTER_DIR, os.platform() === 'win32' ? 'x2t.exe' : 'x2t');

  if (fs.existsSync(x2tPath)) {
    console.log('Converter already installed at', x2tPath);
    return;
  }

  const url = getDownloadUrl();
  const filename = path.basename(new URL(url).pathname);
  const downloadPath = path.join(__dirname, filename);

  console.log(`Downloading ${filename}...`);
  await download(url, downloadPath);

  console.log('Extracting converter directory...');
  if (os.platform() === 'darwin') {
    extractDmg(downloadPath);
  } else {
    extractZip(downloadPath);
  }

  console.log('Cleaning up...');
  cleanup();
  fs.unlinkSync(downloadPath);

  console.log('Done!');
}

main().catch((err) => {
  console.error('Error:', err.message);
  process.exit(1);
});
