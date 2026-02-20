/**
 * Test URL scheme detection regex used in SDK DesktopOfflineAppDocumentEndLoad
 *
 * This tests the fix for Cell/Slide SDKs incorrectly prepending file:// to HTTP URLs.
 * The regex /^[a-zA-Z][a-zA-Z0-9+.-]*:/ detects any valid URL scheme.
 *
 * Run: node test-url-scheme.js
 */

const schemeRegex = /^[a-zA-Z][a-zA-Z0-9+.-]*:/;

let passed = 0;
let failed = 0;

function test(description, condition) {
  if (condition) {
    console.log(`  ✓ ${description}`);
    passed++;
  } else {
    console.error(`  ✗ ${description}`);
    failed++;
  }
}

console.log('\n=== URL Scheme Detection Tests ===\n');

console.log('URLs WITH scheme (should match - no file:// prepended):');
test('http://localhost:38123/api/doc-base/abc', schemeRegex.test('http://localhost:38123/api/doc-base/abc'));
test('https://example.com/path', schemeRegex.test('https://example.com/path'));
test('file:///path/to/document', schemeRegex.test('file:///path/to/document'));
test('file://localhost/path', schemeRegex.test('file://localhost/path'));
test('ftp://server/file', schemeRegex.test('ftp://server/file'));
test('data:image/png;base64,abc', schemeRegex.test('data:image/png;base64,abc'));
test('blob:http://localhost/uuid', schemeRegex.test('blob:http://localhost/uuid'));

console.log('\nPaths WITHOUT scheme (should NOT match - file:// would be prepended):');
test('/absolute/path/to/file', !schemeRegex.test('/absolute/path/to/file'));
test('relative/path/to/file', !schemeRegex.test('relative/path/to/file'));
test('./relative/path', !schemeRegex.test('./relative/path'));
test('../parent/path', !schemeRegex.test('../parent/path'));
test('C:\\Windows\\path (C: looks like scheme)', schemeRegex.test('C:\\Windows\\path')); // C: matches as scheme-like

console.log('\nEdge cases:');
test('Empty string', !schemeRegex.test(''));
test('Just colon :', !schemeRegex.test(':'));
test('Number prefix 1http:', !schemeRegex.test('1http://bad'));
test('Scheme with numbers h2c:', schemeRegex.test('h2c://example')); // Valid per RFC 3986

console.log('\n=== Results ===');
console.log(`Passed: ${passed}`);
console.log(`Failed: ${failed}`);
console.log('================\n');

process.exit(failed > 0 ? 1 : 0);
