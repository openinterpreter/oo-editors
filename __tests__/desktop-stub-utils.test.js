import { describe, test, expect } from 'bun:test';
import {
  parseSpriteScale,
  detectLanguageCode,
  shouldUseEastAsiaVariant,
  getSpriteOptions,
  collectFontNames,
  isImageFile,
  extractMediaFilename,
  buildMediaUrl,
  extractBlobUrl,
  FONT_SPRITE_BASE_WIDTH,
  FONT_SPRITE_ROW_HEIGHT
} from '../editors/desktop-stub-utils.js';

describe('parseSpriteScale', () => {
  test('returns 1 for undefined/null input', () => {
    expect(parseSpriteScale(undefined)).toBe(1);
    expect(parseSpriteScale(null)).toBe(1);
  });

  test('parses @Nx format strings', () => {
    expect(parseSpriteScale('@1x')).toBe(1);
    expect(parseSpriteScale('@2x')).toBe(2);
    expect(parseSpriteScale('@1.5x')).toBe(1.5);
    expect(parseSpriteScale('@3x')).toBe(3);
  });

  test('enforces minimum 0.5 for @Nx format', () => {
    expect(parseSpriteScale('@0.3x')).toBe(0.5);
    expect(parseSpriteScale('@0.1x')).toBe(0.5);
  });

  test('returns 1 for invalid string formats', () => {
    expect(parseSpriteScale('invalid')).toBe(1);
    expect(parseSpriteScale('2x')).toBe(1); // missing @
    expect(parseSpriteScale('@x')).toBe(1); // missing number
  });

  test('handles numeric input', () => {
    expect(parseSpriteScale(2)).toBe(2);
    expect(parseSpriteScale(1.5)).toBe(1.5);
    expect(parseSpriteScale(0.3)).toBe(0.5); // enforces minimum
  });

  test('handles boolean input', () => {
    expect(parseSpriteScale(true)).toBe(2);
    expect(parseSpriteScale(false)).toBe(1);
  });
});

describe('detectLanguageCode', () => {
  test('returns empty string for no location', () => {
    expect(detectLanguageCode(null)).toBe('');
    expect(detectLanguageCode(undefined)).toBe('');
    expect(detectLanguageCode({})).toBe('');
  });

  test('extracts lang from query string', () => {
    expect(detectLanguageCode({ search: '?lang=en' })).toBe('en');
    expect(detectLanguageCode({ search: '?lang=zh-CN' })).toBe('zh-cn');
    expect(detectLanguageCode({ search: '?foo=bar&lang=ja' })).toBe('ja');
  });

  test('handles URL-encoded values', () => {
    expect(detectLanguageCode({ search: '?lang=zh%2DCN' })).toBe('zh-cn');
  });

  test('returns empty for missing lang param', () => {
    expect(detectLanguageCode({ search: '?foo=bar' })).toBe('');
    expect(detectLanguageCode({ search: '' })).toBe('');
  });
});

describe('shouldUseEastAsiaVariant', () => {
  test('returns false for empty/null lang', () => {
    expect(shouldUseEastAsiaVariant('')).toBe(false);
    expect(shouldUseEastAsiaVariant(null)).toBe(false);
    expect(shouldUseEastAsiaVariant(undefined)).toBe(false);
  });

  test('returns true for Chinese', () => {
    expect(shouldUseEastAsiaVariant('zh')).toBe(true);
    expect(shouldUseEastAsiaVariant('zh-cn')).toBe(true);
    expect(shouldUseEastAsiaVariant('zh-tw')).toBe(true);
    expect(shouldUseEastAsiaVariant('zh_CN')).toBe(true);
  });

  test('returns true for Japanese', () => {
    expect(shouldUseEastAsiaVariant('ja')).toBe(true);
    expect(shouldUseEastAsiaVariant('ja-jp')).toBe(true);
  });

  test('returns true for Korean', () => {
    expect(shouldUseEastAsiaVariant('ko')).toBe(true);
    expect(shouldUseEastAsiaVariant('ko-kr')).toBe(true);
  });

  test('returns false for other languages', () => {
    expect(shouldUseEastAsiaVariant('en')).toBe(false);
    expect(shouldUseEastAsiaVariant('en-us')).toBe(false);
    expect(shouldUseEastAsiaVariant('de')).toBe(false);
    expect(shouldUseEastAsiaVariant('fr')).toBe(false);
    expect(shouldUseEastAsiaVariant('ru')).toBe(false);
  });
});

describe('getSpriteOptions', () => {
  test('returns defaults for empty args', () => {
    const result = getSpriteOptions([], '');
    expect(result.scale).toBe(1);
    expect(result.useEA).toBe(false);
  });

  test('detects scale from args', () => {
    expect(getSpriteOptions(['@2x'], '').scale).toBe(2);
    expect(getSpriteOptions([1.5], '').scale).toBe(1.5);
  });

  test('detects EA from _ea suffix in args', () => {
    expect(getSpriteOptions(['fonts_ea'], '').useEA).toBe(true);
    expect(getSpriteOptions(['sprite_ea.png'], '').useEA).toBe(true);
  });

  test('uses language detection for EA', () => {
    expect(getSpriteOptions([], 'zh').useEA).toBe(true);
    expect(getSpriteOptions([], 'ja').useEA).toBe(true);
    expect(getSpriteOptions([], 'en').useEA).toBe(false);
  });

  test('combines scale and EA options', () => {
    const result = getSpriteOptions(['@2x', 'test_ea'], '');
    expect(result.scale).toBe(2);
    expect(result.useEA).toBe(true);
  });
});

describe('collectFontNames', () => {
  test('returns empty array for invalid input', () => {
    expect(collectFontNames(null, false)).toEqual([]);
    expect(collectFontNames(undefined, false)).toEqual([]);
    expect(collectFontNames([], false)).toEqual([]);
    expect(collectFontNames('not an array', false)).toEqual([]);
  });

  test('extracts font names from array', () => {
    const fontsInfos = [
      ['Arial', 'data1', 'data2'],
      ['Times New Roman', 'data1', 'data2'],
      ['Courier', 'data1', 'data2']
    ];
    expect(collectFontNames(fontsInfos, false)).toEqual([
      'Arial',
      'Times New Roman',
      'Courier'
    ]);
  });

  test('skips invalid entries', () => {
    const fontsInfos = [
      ['Arial'],
      null,
      ['Times New Roman'],
      [],
      [null],
      ['Courier']
    ];
    expect(collectFontNames(fontsInfos, false)).toEqual([
      'Arial',
      'Times New Roman',
      'Courier'
    ]);
  });

  test('uses EA variant name when available and requested', () => {
    // Entry format: [name, ..., eaName at index 9]
    const fontsInfos = [
      ['Arial', 1, 2, 3, 4, 5, 6, 7, 8, 'Arial EA'],
      ['Times', 1, 2, 3, 4, 5, 6, 7, 8, ''], // empty EA name
      ['Courier', 1, 2, 3, 4, 5, 6, 7, 8, 'Courier EA']
    ];
    expect(collectFontNames(fontsInfos, true)).toEqual([
      'Arial EA',
      'Times',
      'Courier EA'
    ]);
  });

  test('uses primary name when EA not requested', () => {
    const fontsInfos = [
      ['Arial', 1, 2, 3, 4, 5, 6, 7, 8, 'Arial EA']
    ];
    expect(collectFontNames(fontsInfos, false)).toEqual(['Arial']);
  });
});

describe('isImageFile', () => {
  test('returns false for null/empty', () => {
    expect(isImageFile(null)).toBe(false);
    expect(isImageFile('')).toBe(false);
    expect(isImageFile(undefined)).toBe(false);
  });

  test('recognizes common image extensions', () => {
    expect(isImageFile('photo.jpg')).toBe(true);
    expect(isImageFile('photo.jpeg')).toBe(true);
    expect(isImageFile('image.png')).toBe(true);
    expect(isImageFile('animation.gif')).toBe(true);
    expect(isImageFile('bitmap.bmp')).toBe(true);
    expect(isImageFile('vector.svg')).toBe(true);
    expect(isImageFile('modern.webp')).toBe(true);
  });

  test('handles uppercase extensions', () => {
    expect(isImageFile('photo.JPG')).toBe(true);
    expect(isImageFile('image.PNG')).toBe(true);
  });

  test('handles paths with directories', () => {
    expect(isImageFile('/path/to/photo.jpg')).toBe(true);
    expect(isImageFile('media/image.png')).toBe(true);
  });

  test('returns false for non-image files', () => {
    expect(isImageFile('document.pdf')).toBe(false);
    expect(isImageFile('spreadsheet.xlsx')).toBe(false);
    expect(isImageFile('script.js')).toBe(false);
    expect(isImageFile('noextension')).toBe(false);
  });
});

describe('extractMediaFilename', () => {
  test('returns empty for null/empty', () => {
    expect(extractMediaFilename(null)).toBe('');
    expect(extractMediaFilename('')).toBe('');
    expect(extractMediaFilename(undefined)).toBe('');
  });

  test('strips media/ prefix', () => {
    expect(extractMediaFilename('media/image1.png')).toBe('image1.png');
  });

  test('strips ./media/ prefix', () => {
    expect(extractMediaFilename('./media/image1.png')).toBe('image1.png');
  });

  test('extracts from Editor.bin/media/ path', () => {
    expect(extractMediaFilename('Editor.bin/media/image1.png')).toBe('image1.png');
  });

  test('extracts from arbitrary /media/ path', () => {
    expect(extractMediaFilename('/some/path/media/image1.png')).toBe('image1.png');
  });

  test('returns unchanged if no media/ pattern', () => {
    expect(extractMediaFilename('image1.png')).toBe('image1.png');
    expect(extractMediaFilename('/other/path/image1.png')).toBe('/other/path/image1.png');
  });
});

describe('buildMediaUrl', () => {
  test('constructs correct URL', () => {
    const result = buildMediaUrl('http://localhost:38123', 'abc123', 'image.png');
    expect(result).toBe('http://localhost:38123/api/media/abc123/image.png');
  });

  test('encodes filename', () => {
    const result = buildMediaUrl('http://localhost:38123', 'abc123', 'my image.png');
    expect(result).toBe('http://localhost:38123/api/media/abc123/my%20image.png');
  });

  test('handles special characters', () => {
    const result = buildMediaUrl('http://localhost:38123', 'abc123', 'image (1).png');
    expect(result).toBe('http://localhost:38123/api/media/abc123/image%20(1).png');
  });
});

describe('extractBlobUrl', () => {
  test('returns null for null/empty', () => {
    expect(extractBlobUrl(null)).toBe(null);
    expect(extractBlobUrl('')).toBe(null);
    expect(extractBlobUrl(undefined)).toBe(null);
  });

  test('extracts blob URL from string', () => {
    const blobUrl = 'blob:http://localhost:38123/550e8400-e29b-41d4-a716-446655440000';
    expect(extractBlobUrl(blobUrl)).toBe(blobUrl);
  });

  test('extracts blob URL from wrapped path', () => {
    const wrapped = 'file:///path/to/blob:http://localhost:38123/550e8400-e29b-41d4-a716-446655440000/media/image.png';
    expect(extractBlobUrl(wrapped)).toBe('blob:http://localhost:38123/550e8400-e29b-41d4-a716-446655440000');
  });

  test('returns null for non-blob paths', () => {
    expect(extractBlobUrl('http://localhost:38123/image.png')).toBe(null);
    expect(extractBlobUrl('file:///path/to/image.png')).toBe(null);
    expect(extractBlobUrl('/media/image.png')).toBe(null);
  });
});

describe('constants', () => {
  test('FONT_SPRITE_BASE_WIDTH is defined', () => {
    expect(FONT_SPRITE_BASE_WIDTH).toBe(300);
  });

  test('FONT_SPRITE_ROW_HEIGHT is defined', () => {
    expect(FONT_SPRITE_ROW_HEIGHT).toBe(28);
  });
});

describe('APP_VERSION', () => {
  test('desktop-stub.js APP_VERSION matches package.json version', () => {
    const fs = require('fs');
    const path = require('path');
    const pkg = JSON.parse(fs.readFileSync(path.join(__dirname, '..', 'package.json'), 'utf8'));
    const stub = fs.readFileSync(path.join(__dirname, '..', 'editors', 'desktop-stub.js'), 'utf8');
    const match = stub.match(/var APP_VERSION = '([^']+)'/);
    expect(match).not.toBeNull();
    expect(match[1]).toBe(pkg.version);
  });
});
