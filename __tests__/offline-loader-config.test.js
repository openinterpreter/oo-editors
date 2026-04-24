import { describe, expect, test } from 'bun:test';
import fs from 'fs';
import path from 'path';

const LOADER_PATH = path.resolve(import.meta.dir, '..', 'editors', 'offline-loader-proper.html');

describe('offline loader editorConfig', () => {
  test('disables ONLYOFFICE feature tips', () => {
    const loaderHtml = fs.readFileSync(LOADER_PATH, 'utf8');
    expect(loaderHtml).toMatch(/featuresTips\s*:\s*false/);
  });

  test('streams live editor selections to the host window', () => {
    const loaderHtml = fs.readFileSync(LOADER_PATH, 'utf8');

    expect(loaderHtml).toContain('createSelectionStream');
    expect(loaderHtml).toContain('window.parent');
    expect(loaderHtml).toContain('selectionStream.emitCell(cellInfo)');
    expect(loaderHtml).toContain("asc_registerCallback('asc_onSelectionChanged'");
    expect(loaderHtml).toContain("asc_registerCallback('asc_onSelectionNameChanged'");
    expect(loaderHtml).toContain("asc_registerCallback('asc_onSelectionEnd'");
    expect(loaderHtml).toContain("asc_registerCallback('asc_onCursorMove'");
    expect(loaderHtml).toContain("asc_registerCallback('asc_onFocusObject'");
    expect(loaderHtml).toContain('selectionStream.emitDocument(focusedElements)');
  });
});
