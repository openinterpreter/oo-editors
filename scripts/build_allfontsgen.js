#!/usr/bin/env node

/**
 * Compile the AllFontsGen utility for the current platform.
 *
 * macOS  : produces converter/allfontsgen
 * Windows: produces converter/allfontsgen.exe
 * Linux  : produces converter/allfontsgen (planned, not implemented)
 *
 * Note:
 *   - Only the macOS path is implemented here, assuming converter/
 *     already contains graphics/kernel/UnicodeConverter frameworks.
 *   - Windows/Linux code paths are placeholders—extend as needed.
 */

const { spawnSync } = require('child_process');
const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..');
const CONVERTER_DIR = path.join(ROOT, 'converter');
const DESKTOP_EDITORS_DIR = path.join(ROOT, 'DesktopEditors');
const CORE_DIR = path.join(DESKTOP_EDITORS_DIR, 'core');
const CORE_REPO = 'https://github.com/ONLYOFFICE/core.git';
const CORE_COMMIT = '82e281cf6bf89498e4de6018423b36576706c2b6';
const SRC = path.join(CORE_DIR, 'DesktopEditor', 'AllFontsGen', 'main.cpp');
const WINDOWS_COMPONENTS = ['graphics', 'kernel', 'UnicodeConverter'];
const FORCE_ALLFONTSGEN_REBUILD = process.env.ALLFONTSGEN_FORCE_REBUILD === '1';
const ALLFONTSGEN_DIAGNOSTICS_PATCH = path.join(
  ROOT,
  'scripts',
  'patches',
  'allfontsgen-control-diagnostics.patch'
);
let cachedConverterArch = null;

const platform = process.platform; // 'darwin', 'win32', 'linux'

// Early return if allfontsgen binary already exists
function checkExistingBinary() {
  if (FORCE_ALLFONTSGEN_REBUILD) {
    console.log('[build_allfontsgen] ALLFONTSGEN_FORCE_REBUILD=1, rebuilding binary');
    return false;
  }

  if (platform === 'darwin') {
    const output = path.join(CONVERTER_DIR, 'allfontsgen');
    if (fs.existsSync(output)) {
      console.log('[build_allfontsgen] Binary already exists at', output);
      return true;
    }
  } else if (platform === 'win32') {
    const output = path.join(CONVERTER_DIR, 'allfontsgen.exe');
    if (fs.existsSync(output)) {
      console.log('[build_allfontsgen] Binary already exists at', output);
      return true;
    }
  } else if (platform === 'linux') {
    const output = path.join(CONVERTER_DIR, 'allfontsgen');
    if (fs.existsSync(output)) {
      console.log('[build_allfontsgen] Binary already exists at', output);
      return true;
    }
  }
  return false;
}

if (checkExistingBinary()) {
  process.exit(0);
}

function run(cmd, args, opts = {}) {
  const { status, error } = spawnSync(cmd, args, { stdio: 'inherit', ...opts });
  if (status !== 0) {
    console.error(`[build_allfontsgen] ${cmd} failed`);
    if (error) console.error(error);
    process.exit(status || 1);
  }
}

function cleanup() {
  if (fs.existsSync(DESKTOP_EDITORS_DIR)) {
    console.log('[build_allfontsgen] Cleaning up DesktopEditors directory...');
    fs.rmSync(DESKTOP_EDITORS_DIR, { recursive: true, force: true });
  }
}

function ensureConverterDlls() {
  const missing = WINDOWS_COMPONENTS
    .map((name) => `${name}.dll`)
    .filter((dll) => !fs.existsSync(path.join(CONVERTER_DIR, dll)));

  if (missing.length) {
    console.error('[build_allfontsgen] Missing converter DLLs:', missing.join(', '));
    console.error('[build_allfontsgen] Run download-converter.js before building on Windows.');
    process.exit(1);
  }
}

function readPeMachine(dllPath) {
  const fd = fs.openSync(dllPath, 'r');
  try {
    const dosHeader = Buffer.alloc(64);
    fs.readSync(fd, dosHeader, 0, 64, 0);
    const peOffset = dosHeader.readUInt32LE(0x3c);
    const peHeader = Buffer.alloc(6);
    fs.readSync(fd, peHeader, 0, 6, peOffset);
    const signature = peHeader.readUInt32LE(0);
    if (signature !== 0x4550) {
      throw new Error('Invalid PE signature');
    }
    const machine = peHeader.readUInt16LE(4);
    return machine;
  } finally {
    fs.closeSync(fd);
  }
}

function getConverterArchitecture() {
  if (cachedConverterArch) {
    return cachedConverterArch;
  }

  const kernelDll = path.join(CONVERTER_DIR, 'kernel.dll');
  if (!fs.existsSync(kernelDll)) {
    console.error('[build_allfontsgen] Cannot detect converter architecture. Missing kernel.dll');
    process.exit(1);
  }

  let machine;
  try {
    machine = readPeMachine(kernelDll);
  } catch (err) {
    console.error('[build_allfontsgen] Failed to read PE header from', kernelDll);
    if (err) console.error(err);
    process.exit(1);
  }

  if (machine === 0x8664) {
    cachedConverterArch = 'x64';
  } else if (machine === 0x14c) {
    cachedConverterArch = 'x86';
  } else {
    console.error('[build_allfontsgen] Unsupported converter architecture (machine code:', machine.toString(16), ')');
    process.exit(1);
  }

  console.log('[build_allfontsgen] Detected converter architecture:', cachedConverterArch);
  return cachedConverterArch;
}

function locateVsDevCmd() {
  const programFilesX86 = process.env['ProgramFiles(x86)'];
  const programFiles = programFilesX86 || process.env.ProgramFiles;

  if (!programFiles) {
    console.error('[build_allfontsgen] Unable to locate Program Files. Install Visual Studio Build Tools.');
    process.exit(1);
  }

  const vswherePath = path.join(programFiles, 'Microsoft Visual Studio', 'Installer', 'vswhere.exe');
  if (!fs.existsSync(vswherePath)) {
    console.error('[build_allfontsgen] vswhere.exe not found. Install Visual Studio 2017+ or Build Tools.');
    process.exit(1);
  }

  const result = spawnSync(
    vswherePath,
    ['-latest', '-products', '*', '-requires', 'Microsoft.Component.MSBuild', '-property', 'installationPath'],
    { encoding: 'utf8', stdio: ['ignore', 'pipe', 'inherit'] }
  );

  if (result.status !== 0) {
    console.error('[build_allfontsgen] vswhere.exe failed to locate Visual Studio.');
    process.exit(result.status || 1);
  }

  const installPath = result.stdout.trim().split(/\r?\n/).filter(Boolean).pop();
  if (!installPath) {
    console.error('[build_allfontsgen] Visual Studio installation not found. Ensure Build Tools with MSVC are installed.');
    process.exit(1);
  }

  const devCmd = path.join(installPath, 'Common7', 'Tools', 'VsDevCmd.bat');
  if (!fs.existsSync(devCmd)) {
    console.error('[build_allfontsgen] VsDevCmd.bat not found inside Visual Studio installation.');
    process.exit(1);
  }

  console.log('[build_allfontsgen] Using Visual Studio installation at', installPath);
  return devCmd;
}

function quoteArg(arg) {
  if (/^[A-Za-z0-9._:\\/=+-]+$/.test(arg)) {
    return arg;
  }
  return `"${arg.replace(/"/g, '""')}"`;
}

function runVsDevCommand(vsDevCmd, targetArch, args, opts = {}) {
  if (!Array.isArray(args) || args.length === 0) {
    throw new Error('runVsDevCommand requires a non-empty args array');
  }

  const commandString = args.map(quoteArg).join(' ');
  const hostArch = process.arch === 'x64' ? 'x64' : 'x86';
  const invocation = `"${vsDevCmd}" -arch=${targetArch} -host_arch=${hostArch} && ${commandString}`;
  const spawnOpts = { ...opts };

  console.log('[build_allfontsgen] Running:', invocation);

  if (!('stdio' in spawnOpts)) {
    spawnOpts.stdio = 'inherit';
  }

  if (spawnOpts.stdio !== 'inherit' && !('encoding' in spawnOpts)) {
    spawnOpts.encoding = 'utf8';
  }

  const result = spawnSync('cmd.exe', ['/s', '/c', `"${invocation}"`], {
    ...spawnOpts,
    windowsVerbatimArguments: true,
  });

  if (result.status !== 0) {
    console.error(`[build_allfontsgen] ${args[0]} failed`);
    if (result.error) {
      console.error(result.error);
    }
    process.exit(result.status || 1);
  }

  return result;
}

function parseDumpbinExports(output) {
  const lines = output.split(/\r?\n/);
  const symbols = [];
  let inTable = false;

  for (const line of lines) {
    if (!inTable) {
      if (/ordinal\s+hint\s+RVA\s+name/i.test(line)) {
        inTable = true;
      }
      continue;
    }

    if (/^\s*Summary/.test(line)) {
      break;
    }

    const match = line.match(/^\s*\d+\s+[0-9A-Fa-f]+\s+[0-9A-Fa-f]+\s+(\S+)/);
    if (match) {
      const name = match[1];
      if (!name.includes('[')) {
        symbols.push(name);
      }
    }
  }

  return [...new Set(symbols)];
}

function ensureImportLib(vsDevCmd, targetArch, component) {
  const dllPath = path.join(CONVERTER_DIR, `${component}.dll`);
  const libPath = path.join(CONVERTER_DIR, `${component}.lib`);

  console.log(`[build_allfontsgen] Generating ${component}.lib from ${path.basename(dllPath)}...`);

  if (fs.existsSync(libPath)) {
    fs.rmSync(libPath, { force: true });
  }
  const expPath = libPath.replace(/\.lib$/i, '.exp');
  if (fs.existsSync(expPath)) {
    fs.rmSync(expPath, { force: true });
  }

  const dumpResult = runVsDevCommand(
    vsDevCmd,
    targetArch,
    ['dumpbin.exe', '/nologo', '/exports', dllPath],
    { stdio: ['ignore', 'pipe', 'inherit'], encoding: 'utf8' }
  );

  const exports = parseDumpbinExports(dumpResult.stdout || '');
  if (!exports.length) {
    console.error(`[build_allfontsgen] No exports found in ${dllPath}`);
    process.exit(1);
  }

  const defPath = path.join(CONVERTER_DIR, `${component}.def`);
  const defContent = `LIBRARY ${component}\nEXPORTS\n${exports.join('\n')}\n`;
  fs.writeFileSync(defPath, defContent);

  try {
    const machineFlag = targetArch === 'x64' ? 'x64' : 'x86';
    runVsDevCommand(
      vsDevCmd,
      targetArch,
      ['lib.exe', '/nologo', `/machine:${machineFlag}`, `/def:${defPath}`, `/out:${libPath}`],
      { cwd: ROOT }
    );
  } finally {
    fs.rmSync(defPath, { force: true });
  }
}

function compileWindowsBinary(vsDevCmd, targetArch) {
  const output = path.join(CONVERTER_DIR, 'allfontsgen.exe');

  const includeDirs = [
    CORE_DIR,
    path.join(CORE_DIR, 'DesktopEditor'),
    path.join(CORE_DIR, 'Common'),
    path.join(CORE_DIR, 'UnicodeConverter'),
    path.join(CORE_DIR, 'DesktopEditor', 'fontengine'),
    path.join(CORE_DIR, 'DesktopEditor', 'common'),
    path.join(CORE_DIR, 'DesktopEditor', 'raster'),
    path.join(CORE_DIR, 'DesktopEditor', 'graphics', 'pro'),
  ];

  const includeArgs = includeDirs.flatMap((dir) => ['/I', dir]);
  const defineArgs = [
    '/DKERNEL_USE_DYNAMIC_LIBRARY',
    '/DGRAPHICS_USE_DYNAMIC_LIBRARY',
    '/DNOMINMAX',
    '/DUNICODE',
    '/D_UNICODE',
  ];

  const systemLibs = ['Advapi32.lib', 'Shell32.lib', 'Gdi32.lib', 'User32.lib'];

  const compileArgs = [
    'cl.exe',
    '/std:c++17',
    '/EHsc',
    '/MD',
    '/nologo',
    ...defineArgs,
    ...includeArgs,
    SRC,
    `/Fe${output}`,
    '/link',
    `/LIBPATH:${CONVERTER_DIR}`,
    'graphics.lib',
    'kernel.lib',
    'UnicodeConverter.lib',
    ...systemLibs,
  ];

  console.log('[build_allfontsgen] compiling Windows binary...');
  runVsDevCommand(vsDevCmd, targetArch, compileArgs, { cwd: ROOT });
  console.log('[build_allfontsgen] Windows binary ready at', output);
}

function ensureCoreSources() {
  if (!fs.existsSync(DESKTOP_EDITORS_DIR)) {
    fs.mkdirSync(DESKTOP_EDITORS_DIR, { recursive: true });
  }

  const gitDir = path.join(CORE_DIR, '.git');

  if (fs.existsSync(CORE_DIR) && !fs.existsSync(gitDir)) {
    console.log('[build_allfontsgen] Existing core directory is not a git checkout; removing...');
    fs.rmSync(CORE_DIR, { recursive: true, force: true });
  }

  if (!fs.existsSync(CORE_DIR)) {
    console.log('[build_allfontsgen] Cloning ONLYOFFICE/core repository...');
    run('git', ['clone', CORE_REPO, CORE_DIR], { cwd: ROOT });
  } else {
    console.log('[build_allfontsgen] Reusing existing ONLYOFFICE/core checkout...');
  }

  console.log(`[build_allfontsgen] Checking out ONLYOFFICE/core commit ${CORE_COMMIT}...`);
  run('git', ['fetch', '--all'], { cwd: CORE_DIR });
  run('git', ['checkout', '--force', CORE_COMMIT], { cwd: CORE_DIR });
  console.log('[build_allfontsgen] Updating ONLYOFFICE/core submodules...');
  run('git', ['submodule', 'update', '--init', '--recursive', '--force'], { cwd: CORE_DIR });
}

function applyAllFontsGenDiagnosticsPatch() {
  if (!fs.existsSync(ALLFONTSGEN_DIAGNOSTICS_PATCH)) {
    console.error('[build_allfontsgen] Missing diagnostics patch:', ALLFONTSGEN_DIAGNOSTICS_PATCH);
    process.exit(1);
  }

  console.log('[build_allfontsgen] Applying AllFontsGen diagnostics patch...');
  run('git', ['apply', '--unidiff-zero', '--whitespace=nowarn', ALLFONTSGEN_DIAGNOSTICS_PATCH], {
    cwd: CORE_DIR,
  });
}

ensureCoreSources();
applyAllFontsGenDiagnosticsPatch();

if (!fs.existsSync(SRC)) {
  console.error('[build_allfontsgen] main.cpp not found:', SRC);
  cleanup();
  process.exit(1);
}

if (platform === 'darwin') {
  if (!fs.existsSync(CONVERTER_DIR)) {
    console.error('[build_allfontsgen] missing converter/ directory with frameworks');
    process.exit(1);
  }

  const output = path.join(CONVERTER_DIR, 'allfontsgen');
  const clangArgs = [
    '-std=c++17',
    SRC,
    '-DKERNEL_USE_DYNAMIC_LIBRARY',
    '-DGRAPHICS_USE_DYNAMIC_LIBRARY',
    '-IDesktopEditors/core',
    '-IDesktopEditors/core/DesktopEditor',
    '-IDesktopEditors/core/Common',
    '-IDesktopEditors/core/UnicodeConverter',
    '-IDesktopEditors/core/DesktopEditor/fontengine',
    '-IDesktopEditors/core/DesktopEditor/common',
    '-IDesktopEditors/core/DesktopEditor/raster',
    '-IDesktopEditors/core/DesktopEditor/graphics/pro',
    `-F${CONVERTER_DIR}`,
    '-framework', 'graphics',
    '-framework', 'kernel',
    '-framework', 'UnicodeConverter',
    `-Wl,-rpath,@executable_path`,
    '-o', output,
  ];

  console.log('[build_allfontsgen] compiling macOS binary…');
  run('clang++', clangArgs, { cwd: ROOT });
  fs.chmodSync(output, 0o755);
  console.log('[build_allfontsgen] macOS binary ready at', output);
  cleanup();
  process.exit(0);
}

if (platform === 'win32') {
  ensureConverterDlls();
  const converterArch = getConverterArchitecture();
  const vsDevCmd = locateVsDevCmd();
  WINDOWS_COMPONENTS.forEach((component) => ensureImportLib(vsDevCmd, converterArch, component));
  compileWindowsBinary(vsDevCmd, converterArch);
  cleanup();
  process.exit(0);
}

if (platform === 'linux') {
  console.error('[build_allfontsgen] Linux build not implemented. Compile with g++ and place result in converter/allfontsgen.');
  process.exit(1);
}

console.error('[build_allfontsgen] Unsupported platform:', platform);
process.exit(1);
