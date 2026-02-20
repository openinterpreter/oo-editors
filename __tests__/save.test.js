import { describe, test, expect, beforeAll, afterAll } from 'bun:test';
import path from 'path';
import fs from 'fs';
import os from 'os';
import crypto from 'crypto';

const SERVER_URL = process.env.SERVER_URL || 'http://localhost:38123';
const FIXTURES_DIR = path.resolve(import.meta.dir, '..', '.github/assets');
const tempFiles = [];

function tempPath() {
  const p = path.join(os.tmpdir(), `oo-editors-test-${crypto.randomUUID()}.csv`);
  tempFiles.push(p);
  return p;
}

beforeAll(async () => {
  try {
    const res = await fetch(`${SERVER_URL}/healthcheck`);
    if (!res.ok) throw new Error(`status ${res.status}`);
  } catch {
    throw new Error(`Server not running at ${SERVER_URL}. Start with: bun run server`);
  }
});

afterAll(() => {
  for (const p of tempFiles) {
    try { fs.unlinkSync(p); } catch {}
  }
});

async function convertAndSave(fixtureName) {
  const fixturePath = path.join(FIXTURES_DIR, fixtureName);

  const convertRes = await fetch(
    `${SERVER_URL}/api/convert?filepath=${encodeURIComponent(fixturePath)}`
  );
  expect(convertRes.status).toBe(200);

  const fileHash = convertRes.headers.get('x-file-hash');
  expect(fileHash).toBeTruthy();

  const binary = await convertRes.arrayBuffer();
  expect(binary.byteLength).toBeGreaterThan(0);

  const outPath = tempPath();
  const saveRes = await fetch(
    `${SERVER_URL}/api/save?filepath=${encodeURIComponent(outPath)}&filehash=${fileHash}`,
    {
      method: 'POST',
      headers: { 'Content-Type': 'application/octet-stream' },
      body: binary,
    }
  );
  expect(saveRes.status).toBe(200);

  const json = await saveRes.json();
  expect(json.success).toBe(true);
}

describe('CSV save round-trip', () => {
  test('should save simple.csv', async () => {
    await convertAndSave('simple.csv');
  });

  test('should save hard.csv', async () => {
    await convertAndSave('hard.csv');
  }, 60_000);
});
