import { describe, test, expect } from 'bun:test';
import {
  getX2TFormatCode,
  getOutputFormatInfo,
  generateFileHash,
  getDocTypeFromFilename,
  isAbsolutePath,
  getContentType,
  generateX2TConfig,
  extractFilePathFromUrl,
  isXLSXSignature
} from '../server-utils.js';

describe('getX2TFormatCode', () => {
  test('returns correct code for spreadsheet formats', () => {
    expect(getX2TFormatCode('xlsx')).toBe(257);
    expect(getX2TFormatCode('xls')).toBe(257);
    expect(getX2TFormatCode('ods')).toBe(257);
    expect(getX2TFormatCode('csv')).toBe(260);
  });

  test('returns correct code for document formats', () => {
    expect(getX2TFormatCode('docx')).toBe(65);
    expect(getX2TFormatCode('doc')).toBe(65);
    expect(getX2TFormatCode('odt')).toBe(65);
    expect(getX2TFormatCode('txt')).toBe(65);
    expect(getX2TFormatCode('rtf')).toBe(65);
    expect(getX2TFormatCode('html')).toBe(65);
  });

  test('returns correct code for presentation formats', () => {
    expect(getX2TFormatCode('pptx')).toBe(129);
    expect(getX2TFormatCode('ppt')).toBe(129);
    expect(getX2TFormatCode('odp')).toBe(129);
  });

  test('returns correct code for PDF', () => {
    expect(getX2TFormatCode('pdf')).toBe(513);
  });

  test('handles uppercase input', () => {
    expect(getX2TFormatCode('XLSX')).toBe(257);
    expect(getX2TFormatCode('DOCX')).toBe(65);
    expect(getX2TFormatCode('PDF')).toBe(513);
  });

  test('returns null for unsupported formats', () => {
    expect(getX2TFormatCode('unknown')).toBe(null);
    expect(getX2TFormatCode('zip')).toBe(null);
    expect(getX2TFormatCode('mp3')).toBe(null);
  });
});

describe('getOutputFormatInfo', () => {
  test('returns info for spreadsheet extensions', () => {
    expect(getOutputFormatInfo('.xlsx')).toEqual({ code: 257, name: 'XLSX' });
    expect(getOutputFormatInfo('.xls')).toEqual({ code: 257, name: 'XLSX' });
    expect(getOutputFormatInfo('.ods')).toEqual({ code: 257, name: 'XLSX' });
  });

  test('returns info for CSV', () => {
    expect(getOutputFormatInfo('.csv')).toEqual({ code: 260, name: 'CSV' });
  });

  test('returns info for document extensions', () => {
    expect(getOutputFormatInfo('.docx')).toEqual({ code: 65, name: 'DOCX' });
    expect(getOutputFormatInfo('.doc')).toEqual({ code: 65, name: 'DOCX' });
    expect(getOutputFormatInfo('.odt')).toEqual({ code: 65, name: 'DOCX' });
    expect(getOutputFormatInfo('.txt')).toEqual({ code: 65, name: 'DOCX' });
    expect(getOutputFormatInfo('.rtf')).toEqual({ code: 65, name: 'DOCX' });
    expect(getOutputFormatInfo('.html')).toEqual({ code: 65, name: 'DOCX' });
  });

  test('returns info for presentation extensions', () => {
    expect(getOutputFormatInfo('.pptx')).toEqual({ code: 129, name: 'PPTX' });
    expect(getOutputFormatInfo('.ppt')).toEqual({ code: 129, name: 'PPTX' });
    expect(getOutputFormatInfo('.odp')).toEqual({ code: 129, name: 'PPTX' });
  });

  test('handles uppercase', () => {
    expect(getOutputFormatInfo('.XLSX')).toEqual({ code: 257, name: 'XLSX' });
    expect(getOutputFormatInfo('.CSV')).toEqual({ code: 260, name: 'CSV' });
  });

  test('returns null for unsupported', () => {
    expect(getOutputFormatInfo('.pdf')).toBe(null);
    expect(getOutputFormatInfo('.zip')).toBe(null);
  });
});

describe('generateFileHash', () => {
  test('generates consistent MD5 hash', () => {
    const hash1 = generateFileHash('/path/to/file.xlsx');
    const hash2 = generateFileHash('/path/to/file.xlsx');
    expect(hash1).toBe(hash2);
  });

  test('generates different hashes for different paths', () => {
    const hash1 = generateFileHash('/path/to/file1.xlsx');
    const hash2 = generateFileHash('/path/to/file2.xlsx');
    expect(hash1).not.toBe(hash2);
  });

  test('returns 32-character hex string', () => {
    const hash = generateFileHash('/any/path.xlsx');
    expect(hash).toMatch(/^[a-f0-9]{32}$/);
  });
});

describe('getDocTypeFromFilename', () => {
  test('returns cell for spreadsheet files', () => {
    expect(getDocTypeFromFilename('test.xlsx')).toBe('cell');
    expect(getDocTypeFromFilename('test.xls')).toBe('cell');
    expect(getDocTypeFromFilename('test.ods')).toBe('cell');
    expect(getDocTypeFromFilename('test.csv')).toBe('cell');
  });

  test('returns word for document files', () => {
    expect(getDocTypeFromFilename('test.docx')).toBe('word');
    expect(getDocTypeFromFilename('test.doc')).toBe('word');
    expect(getDocTypeFromFilename('test.odt')).toBe('word');
    expect(getDocTypeFromFilename('test.txt')).toBe('word');
    expect(getDocTypeFromFilename('test.rtf')).toBe('word');
    expect(getDocTypeFromFilename('test.html')).toBe('word');
  });

  test('returns slide for presentation files', () => {
    expect(getDocTypeFromFilename('test.pptx')).toBe('slide');
    expect(getDocTypeFromFilename('test.ppt')).toBe('slide');
    expect(getDocTypeFromFilename('test.odp')).toBe('slide');
  });

  test('returns slide for unknown extensions (default)', () => {
    expect(getDocTypeFromFilename('unknown.xyz')).toBe('slide');
  });

  test('handles paths with directories', () => {
    expect(getDocTypeFromFilename('/path/to/test.xlsx')).toBe('cell');
    expect(getDocTypeFromFilename('folder/test.docx')).toBe('word');
  });
});

describe('isAbsolutePath', () => {
  test('recognizes POSIX absolute paths', () => {
    expect(isAbsolutePath('/Users/test/file.xlsx')).toBe(true);
    expect(isAbsolutePath('/home/user/doc.docx')).toBe(true);
    expect(isAbsolutePath('/')).toBe(true);
  });

  test('recognizes Windows absolute paths', () => {
    expect(isAbsolutePath('C:\\Users\\test\\file.xlsx')).toBe(true);
    expect(isAbsolutePath('D:\\Documents\\doc.docx')).toBe(true);
  });

  test('rejects relative paths', () => {
    expect(isAbsolutePath('file.xlsx')).toBe(false);
    expect(isAbsolutePath('./file.xlsx')).toBe(false);
    expect(isAbsolutePath('../file.xlsx')).toBe(false);
    expect(isAbsolutePath('folder/file.xlsx')).toBe(false);
  });
});

describe('getContentType', () => {
  test('returns correct type for images', () => {
    expect(getContentType('.png')).toBe('image/png');
    expect(getContentType('.jpg')).toBe('image/jpeg');
    expect(getContentType('.jpeg')).toBe('image/jpeg');
    expect(getContentType('.gif')).toBe('image/gif');
    expect(getContentType('.svg')).toBe('image/svg+xml');
    expect(getContentType('.webp')).toBe('image/webp');
  });

  test('returns correct type for fonts', () => {
    expect(getContentType('.ttf')).toBe('font/ttf');
    expect(getContentType('.otf')).toBe('font/otf');
    expect(getContentType('.woff')).toBe('font/woff');
    expect(getContentType('.woff2')).toBe('font/woff2');
  });

  test('returns correct type for documents', () => {
    expect(getContentType('.pdf')).toBe('application/pdf');
    expect(getContentType('.xlsx')).toBe('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    expect(getContentType('.docx')).toBe('application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    expect(getContentType('.pptx')).toBe('application/vnd.openxmlformats-officedocument.presentationml.presentation');
  });

  test('handles uppercase', () => {
    expect(getContentType('.PNG')).toBe('image/png');
    expect(getContentType('.PDF')).toBe('application/pdf');
  });

  test('returns octet-stream for unknown', () => {
    expect(getContentType('.xyz')).toBe('application/octet-stream');
    expect(getContentType('.unknown')).toBe('application/octet-stream');
  });
});

describe('generateX2TConfig', () => {
  test('generates valid XML with required fields', () => {
    const xml = generateX2TConfig({
      inputPath: '/input/file.xlsx',
      outputPath: '/output/Editor.bin',
      filename: 'file.xlsx',
      formatTo: 8192,
      fontDir: '/fonts',
      themeDir: '/themes'
    });

    expect(xml).toContain('<?xml version="1.0" encoding="utf-8"?>');
    expect(xml).toContain('<m_sFileFrom>/input/file.xlsx</m_sFileFrom>');
    expect(xml).toContain('<m_sFileTo>/output/Editor.bin</m_sFileTo>');
    expect(xml).toContain('<m_sTitle>file.xlsx</m_sTitle>');
    expect(xml).toContain('<m_nFormatTo>8192</m_nFormatTo>');
    expect(xml).toContain('<m_sFontDir>/fonts</m_sFontDir>');
    expect(xml).toContain('<m_sThemeDir>/themes</m_sThemeDir>');
  });

  test('includes formatFrom when provided', () => {
    const xml = generateX2TConfig({
      inputPath: '/input/file.bin',
      outputPath: '/output/file.xlsx',
      filename: 'file.xlsx',
      formatTo: 257,
      formatFrom: 8192,
      fontDir: '/fonts',
      themeDir: '/themes'
    });

    expect(xml).toContain('<m_nFormatFrom>8192</m_nFormatFrom>');
  });

  test('excludes formatFrom when not provided', () => {
    const xml = generateX2TConfig({
      inputPath: '/input/file.xlsx',
      outputPath: '/output/Editor.bin',
      filename: 'file.xlsx',
      formatTo: 8192,
      fontDir: '/fonts',
      themeDir: '/themes'
    });

    expect(xml).not.toContain('<m_nFormatFrom>');
  });
});

describe('extractFilePathFromUrl', () => {
  test('extracts path from OnlyOffice URL format', () => {
    const url = 'http://localhost:38123/api/onlyoffice/files/Users/test/file.xlsx';
    expect(extractFilePathFromUrl(url)).toBe('/Users/test/file.xlsx');
  });

  test('handles URL-encoded paths', () => {
    const url = 'http://localhost:38123/api/onlyoffice/files/Users/test/my%20file.xlsx';
    expect(extractFilePathFromUrl(url)).toBe('/Users/test/my file.xlsx');
  });

  test('works with https', () => {
    const url = 'https://example.com/api/onlyoffice/files/path/to/doc.docx';
    expect(extractFilePathFromUrl(url)).toBe('/path/to/doc.docx');
  });

  test('returns null for non-matching URLs', () => {
    expect(extractFilePathFromUrl('http://localhost:38123/other/path')).toBe(null);
    expect(extractFilePathFromUrl('http://localhost:38123/api/convert')).toBe(null);
  });

  test('returns null for non-URL input', () => {
    expect(extractFilePathFromUrl('/local/path/file.xlsx')).toBe(null);
    expect(extractFilePathFromUrl('file.xlsx')).toBe(null);
    expect(extractFilePathFromUrl(null)).toBe(null);
    expect(extractFilePathFromUrl('')).toBe(null);
  });
});

describe('isXLSXSignature', () => {
  test('detects XLSX/ZIP signature (PK)', () => {
    const xlsxData = Buffer.from([0x50, 0x4B, 0x03, 0x04]);
    expect(isXLSXSignature(xlsxData)).toBe(true);
  });

  test('rejects non-XLSX data', () => {
    const binaryData = Buffer.from([0x00, 0x00, 0x00, 0x00]);
    expect(isXLSXSignature(binaryData)).toBe(false);
  });

  test('rejects too-short data', () => {
    expect(isXLSXSignature(Buffer.from([0x50]))).toBe(false);
    expect(isXLSXSignature(Buffer.from([]))).toBe(false);
  });

  test('handles null/undefined', () => {
    expect(isXLSXSignature(null)).toBe(false);
    expect(isXLSXSignature(undefined)).toBe(false);
  });
});
