import { describe, expect, test } from 'bun:test';
import fs from 'fs';
import path from 'path';

const OFFLINE_LOADER_PATH = path.resolve(import.meta.dir, '../editors/offline-loader-proper.html');

describe('offline-loader dialog behavior', () => {
  test('does not auto-click generic in-editor dialogs', () => {
    const html = fs.readFileSync(OFFLINE_LOADER_PATH, 'utf8');

    expect(html.includes("setInterval(dismissDialogsOnce, 2000)")).toBe(false);
    expect(html.includes("text === 'got it' || text === 'ok'")).toBe(false);
    expect(html.includes("console.log('[DISMISS] Clicking button:', text)")).toBe(false);
  });

  test('keeps targeted suppression for expected -82 load error', () => {
    const html = fs.readFileSync(OFFLINE_LOADER_PATH, 'utf8');

    expect(html.includes('if (errorCode === -82)')).toBe(true);
    expect(html.includes('[SUPPRESSED] Error -82 is expected during offline loading')).toBe(true);
  });
});
