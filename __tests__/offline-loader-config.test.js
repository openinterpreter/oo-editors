import { describe, expect, test } from 'bun:test';
import fs from 'fs';
import path from 'path';

const LOADER_PATH = path.resolve(import.meta.dir, '..', 'editors', 'offline-loader-proper.html');

describe('offline loader editorConfig', () => {
  test('disables ONLYOFFICE feature tips', () => {
    const loaderHtml = fs.readFileSync(LOADER_PATH, 'utf8');
    expect(loaderHtml).toMatch(/featuresTips\s*:\s*false/);
  });
});
