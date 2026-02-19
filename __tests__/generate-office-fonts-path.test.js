import { afterEach, describe, expect, test } from 'bun:test';
import fs from 'fs';
import os from 'os';
import path from 'path';
import { spawnSync } from 'child_process';

const SCRIPT_PATH = path.resolve(import.meta.dir, '..', 'scripts', 'generate_office_fonts.js');

const tempDirs = [];

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

function runGenerator({ fontDataDir, writeOutputs }) {
  const mockDir = makeTempDir('oo-editors-font-mock-');
  const preloadPath = writeMockPreload(mockDir);

  const env = {
    ...process.env,
    FONT_DATA_DIR: fontDataDir,
    MOCK_EXIT_CODE: '0',
    MOCK_WRITE_OUTPUTS: writeOutputs ? '1' : '0',
    NODE_OPTIONS: [process.env.NODE_OPTIONS, `--require=${preloadPath}`]
      .filter(Boolean)
      .join(' '),
  };

  return spawnSync('node', [SCRIPT_PATH], {
    env,
    encoding: 'utf8',
  });
}

afterEach(() => {
  while (tempDirs.length > 0) {
    const dir = tempDirs.pop();
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

describe('generate_office_fonts path handling', () => {
  test('supports non-ASCII FONT_DATA_DIR paths when outputs are written', () => {
    const root = makeTempDir('oo-editors-fontdata-');
    const fontDataDir = path.join(root, 'C', 'Users', 'دانيال', 'AppData', 'Roaming', 'interpreter', 'office-extension-fontdata');

    const result = runGenerator({ fontDataDir, writeOutputs: true });

    expect(result.status).toBe(0);
    expect(fs.existsSync(path.join(fontDataDir, 'AllFonts.js'))).toBe(true);
    expect(fs.existsSync(path.join(fontDataDir, 'font_selection.bin'))).toBe(true);
  });

  test('fails when generator exits 0 but does not write metadata files', () => {
    const root = makeTempDir('oo-editors-fontdata-');
    const fontDataDir = path.join(root, 'C', 'Users', 'دانيال', 'AppData', 'Roaming', 'interpreter', 'office-extension-fontdata');

    const result = runGenerator({ fontDataDir, writeOutputs: false });

    expect(result.status).toBe(1);
    expect(result.stderr).toContain('font metadata files were not created');
    expect(fs.existsSync(path.join(fontDataDir, 'AllFonts.js'))).toBe(false);
  });
});
