import { afterEach, describe, expect, test } from 'bun:test';
import fs from 'fs';
import os from 'os';
import path from 'path';
import { spawnSync } from 'child_process';

const SCRIPT_PATH = path.resolve(import.meta.dir, '..', 'scripts', 'generate_office_fonts.js');
const CONVERTER_DIR = path.resolve(import.meta.dir, '..', 'converter');
const DUMMY_BIN = path.join(CONVERTER_DIR, process.platform === 'win32' ? 'allfontsgen.exe' : 'allfontsgen');

const tempDirs = [];
let createdDummyBin = false;

function makeTempDir(prefix) {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), prefix));
  tempDirs.push(dir);
  return dir;
}

function writeMockPreload(dir) {
  const preloadPath = path.join(dir, 'mock-spawn.js');
  fs.writeFileSync(
    preloadPath,
    `const childProcess = require('child_process');
const fs = require('fs');
const path = require('path');

childProcess.spawnSync = (_bin, args) => {
  if (process.env.MOCK_ARGS_FILE) {
    fs.writeFileSync(process.env.MOCK_ARGS_FILE, JSON.stringify(args));
  }

  const allFontsArg = args.find((arg) => arg.startsWith('--allfonts='));
  const selectionArg = args.find((arg) => arg.startsWith('--selection='));
  const allFontsPath = allFontsArg ? allFontsArg.slice('--allfonts='.length) : null;
  const selectionPath = selectionArg ? selectionArg.slice('--selection='.length) : null;

  if (process.env.MOCK_WRITE_OUTPUTS === '1' && allFontsPath && selectionPath) {
    fs.mkdirSync(path.dirname(allFontsPath), { recursive: true });
    fs.writeFileSync(allFontsPath, 'window.AllFonts = [];');
    fs.writeFileSync(selectionPath, Buffer.alloc(0));
  }

  return { status: Number(process.env.MOCK_EXIT_CODE || 0) };
};
`
  );

  return preloadPath;
}

function ensureDummyBin() {
  if (!fs.existsSync(DUMMY_BIN)) {
    fs.mkdirSync(CONVERTER_DIR, { recursive: true });
    fs.writeFileSync(DUMMY_BIN, '');
    fs.chmodSync(DUMMY_BIN, 0o755);
    createdDummyBin = true;
  }
}

function runGenerator({ fontDataDir, writeOutputs, additionalFontDirs = [] }) {
  ensureDummyBin();
  const mockDir = makeTempDir('oo-editors-font-mock-');
  const preloadPath = writeMockPreload(mockDir);
  const argsFile = path.join(mockDir, 'spawn-args.json');

  const env = {
    ...process.env,
    FONT_DATA_DIR: fontDataDir,
    ...(additionalFontDirs.length > 0 ? { ADDITIONAL_FONT_DIRS: additionalFontDirs.join(path.delimiter) } : {}),
    MOCK_ARGS_FILE: argsFile,
    MOCK_EXIT_CODE: '0',
    MOCK_WRITE_OUTPUTS: writeOutputs ? '1' : '0',
    NODE_OPTIONS: [process.env.NODE_OPTIONS, `--require=${preloadPath}`]
      .filter(Boolean)
      .join(' '),
  };

  const result = spawnSync('node', [SCRIPT_PATH], {
    env,
    encoding: 'utf8',
    timeout: 10000,
  });

  const parsedArgs = fs.existsSync(argsFile)
    ? JSON.parse(fs.readFileSync(argsFile, 'utf8'))
    : null;

  return { result, args: parsedArgs };
}

afterEach(() => {
  while (tempDirs.length > 0) {
    const dir = tempDirs.pop();
    fs.rmSync(dir, { recursive: true, force: true });
  }
  if (createdDummyBin) {
    fs.rmSync(DUMMY_BIN, { force: true });
    createdDummyBin = false;
  }
});

describe('generate_office_fonts path handling', () => {
  test('should succeed when FONT_DATA_DIR contains non-ASCII characters', () => {
    const root = makeTempDir('oo-editors-fontdata-');
    const fontDataDir = path.join(root, 'دانيال', 'fontdata');

    const { result } = runGenerator({ fontDataDir, writeOutputs: true });

    expect(result.status).toBe(0);
    expect(fs.existsSync(path.join(fontDataDir, 'AllFonts.js'))).toBe(true);
    expect(fs.existsSync(path.join(fontDataDir, 'font_selection.bin'))).toBe(true);
  }, 15000);

  test('should fail when generator exits 0 but does not write metadata files', () => {
    const root = makeTempDir('oo-editors-fontdata-');
    const fontDataDir = path.join(root, 'دانيال', 'fontdata');

    const { result } = runGenerator({ fontDataDir, writeOutputs: false });

    expect(result.status).toBe(1);
    expect(result.stderr).toContain('font metadata files were not created');
    expect(fs.existsSync(path.join(fontDataDir, 'AllFonts.js'))).toBe(false);
  });

  test('should include additional font directories via input symlinks', () => {
    const root = makeTempDir('oo-editors-fontdata-');
    const fontDataDir = path.join(root, 'fontdata');
    const extra1 = path.join(root, 'extra-fonts-1');
    const extra2 = path.join(root, 'extra-fonts-2');
    fs.mkdirSync(extra1, { recursive: true });
    fs.mkdirSync(extra2, { recursive: true });

    const { result, args } = runGenerator({
      fontDataDir,
      writeOutputs: true,
      additionalFontDirs: [extra1, extra2],
    });

    expect(result.status).toBe(0);
    expect(args).toContain(`--input=${path.join(fontDataDir, 'fonts')}`);
    expect(fs.realpathSync(path.join(fontDataDir, 'fonts', 'additional-font-dir-1'))).toBe(fs.realpathSync(extra1));
    expect(fs.realpathSync(path.join(fontDataDir, 'fonts', 'additional-font-dir-2'))).toBe(fs.realpathSync(extra2));
  });

  test('should fail when an additional font directory is missing', () => {
    const root = makeTempDir('oo-editors-fontdata-');
    const fontDataDir = path.join(root, 'fontdata');
    const missing = path.join(root, 'missing-font-dir');

    const { result } = runGenerator({
      fontDataDir,
      writeOutputs: true,
      additionalFontDirs: [missing],
    });

    expect(result.status).toBe(1);
    expect(result.stderr).toContain('additional font directory does not exist');
  });
});
