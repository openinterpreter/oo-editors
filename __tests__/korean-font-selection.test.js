const { describe, test, expect } = require('bun:test');
const fs = require('fs');
const vm = require('vm');
const path = require('path');

// Load the real map.js in a sandbox with minimal stubs.
// CFontSelectList.Init() builds language definitions (our fix target)
// and checkText() runs language detection -- both are self-contained
// once we skip the font_selection.bin binary loading.

function createSandbox() {
  const sandbox = {
    window: {},
    AscFonts: {
      FT_Common: { IntToUInt: (v) => v >>> 0 },
      FontStream: function () {},
      allocate: () => new Uint8Array(0),
      g_font_infos: [],
      g_fonts_streams: [],
      getEmbeddedFontPrefix: () => '__embedded__',
    },
    AscCommon: {
      Base64: { decode: () => new Uint8Array(0) },
      FileStream: function () {},
      g_font_loader: {},
    },
    AscWord: {
      fontslot_ASCII: 0x01,
      fontslot_EastAsia: 0x02,
      fontslot_CS: 0x04,
      fontslot_HAnsi: 0x08,
    },
  };

  // map.js IIFE receives `window` as param and attaches to it
  sandbox.window.AscFonts = sandbox.AscFonts;
  sandbox.window.AscCommon = sandbox.AscCommon;
  sandbox.window.AscWord = sandbox.AscWord;
  sandbox.window['g_fonts_selection_bin'] = '';
  sandbox.window['g_fonts_streams'] = [];

  return sandbox;
}

function loadFontSelectList() {
  const sandbox = createSandbox();
  const mapSource = fs.readFileSync(
    path.join(__dirname, '..', 'sdkjs', 'common', 'libfont', 'map.js'),
    'utf8'
  );

  vm.runInNewContext(mapSource, sandbox);

  const app = sandbox.AscFonts.g_fontApplication;
  const fsl = app.g_fontSelections;
  // Call Init to build language definitions (skips binary loading since g_fonts_selection_bin = "")
  fsl.Init();
  return fsl;
}

let fontSelectList;
try {
  fontSelectList = loadFontSelectList();
} catch (e) {
  console.error('Failed to load map.js:', e.message);
}

describe('Korean language detection (real map.js)', () => {
  test('should load CFontSelectList with 4 languages', () => {
    expect(fontSelectList).toBeTruthy();
    expect(fontSelectList.Languages.length).toBe(4);
  });

  test('Korean language should include CJK Unified Ideographs range', () => {
    // Languages[1] is Korean (index 1, after Arabic at index 0)
    const korean = fontSelectList.Languages[1];
    expect(korean.Type).toBe(2); // LanguagesFontSelectTypes.Korean
    expect(korean.checkChar(0x4e00)).toBe(true); // start of CJK Unified
    expect(korean.checkChar(0x9fff)).toBe(true); // end of CJK Unified
    expect(korean.checkChar(0x5b57)).toBe(true); // U+5B57 "字" (Hanja)
  });

  test('Korean language should include CJK Compatibility Ideographs range', () => {
    const korean = fontSelectList.Languages[1];
    expect(korean.checkChar(0xf900)).toBe(true);
    expect(korean.checkChar(0xfaff)).toBe(true);
  });

  test('Korean language should still include Hangul ranges', () => {
    const korean = fontSelectList.Languages[1];
    expect(korean.checkChar(0x1100)).toBe(true); // Hangul Jamo
    expect(korean.checkChar(0xac00)).toBe(true); // Hangul Syllables start
    expect(korean.checkChar(0xd7af)).toBe(true); // Hangul Syllables end
    expect(korean.checkChar(0x3130)).toBe(true); // Hangul Compatibility Jamo
  });

  test('Japanese should NOT cover Hangul Syllables', () => {
    const japan = fontSelectList.Languages[2];
    expect(japan.Type).toBe(3); // LanguagesFontSelectTypes.Japan
    expect(japan.checkChar(0xac00)).toBe(false);
    expect(japan.checkChar(0xd7af)).toBe(false);
  });

  test('Chinese should NOT cover Hangul Syllables', () => {
    const chinese = fontSelectList.Languages[3];
    expect(chinese.Type).toBe(4); // LanguagesFontSelectTypes.Chinese
    expect(chinese.checkChar(0xac00)).toBe(false);
    expect(chinese.checkChar(0xd7af)).toBe(false);
  });

  test('mixed Hangul + Hanja text should detect as Korean', () => {
    // U+D55C "한" (Hangul) + U+5B57 "字" (Hanja)
    const result = fontSelectList.checkText('\uD55C\u5B57');
    expect(result).toBe(2); // Korean
  });

  test('mixed Hangul + Hanja + English should detect as Korean', () => {
    // "한글abc漢字文化" - more CJK than English so _percent_by_english > 0
    const result = fontSelectList.checkText('\uD55C\uAE00abc\u6F22\u5B57\u6587\u5316');
    expect(result).toBe(2); // Korean
  });

  test('pure Hangul should detect as Korean', () => {
    // "안녕하세요"
    const result = fontSelectList.checkText('\uC548\uB155\uD558\uC138\uC694');
    expect(result).toBe(2); // Korean
  });

  test('pure CJK without Hangul should detect as Chinese (last match wins)', () => {
    // "漢字" - ambiguous, Chinese wins as last matcher
    const result = fontSelectList.checkText('\u6F22\u5B57');
    expect(result).toBe(4); // Chinese
  });

  test('Hiragana + CJK should detect as Japanese', () => {
    // "日本の" - CJK + Hiragana "の" (U+306E)
    const result = fontSelectList.checkText('\u65E5\u672C\u306E');
    expect(result).toBe(3); // Japan
  });
});
