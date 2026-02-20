/**
 * Tests for desktop-stub-utils.js
 * Run: bun test-desktop-stub-utils.js
 */

const utils = require('./editors/desktop-stub-utils.js');

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

console.log('\n=== buildMediaUrl Tests ===\n');

console.log('Valid inputs:');
test('builds URL with all params',
  utils.buildMediaUrl('http://localhost:38123', 'abc123', 'image.png') ===
  'http://localhost:38123/api/media/abc123/image.png');

test('encodes filename with spaces',
  utils.buildMediaUrl('http://localhost:38123', 'abc', 'my image.png').includes('my%20image.png'));

test('encodes unicode filename',
  utils.buildMediaUrl('http://localhost:38123', 'abc', '画像.png').includes('%E7%94%BB%E5%83%8F.png'));

console.log('\nMissing fileHash (BUG FIX - should return null):');
test('returns null when fileHash is undefined',
  utils.buildMediaUrl('http://localhost:38123', undefined, 'image.png') === null);

test('returns null when fileHash is null',
  utils.buildMediaUrl('http://localhost:38123', null, 'image.png') === null);

test('returns null when fileHash is empty string',
  utils.buildMediaUrl('http://localhost:38123', '', 'image.png') === null);

console.log('\n=== Results ===');
console.log(`Passed: ${passed}`);
console.log(`Failed: ${failed}`);
console.log('================\n');

process.exit(failed > 0 ? 1 : 0);
