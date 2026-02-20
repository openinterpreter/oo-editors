const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');
const os = require('os');
const http = require('http');

const SERVER_PORT = Number.parseInt(process.env.SERVER_PORT || '38123', 10);
const SERVER_URL = process.env.SERVER_URL || `http://localhost:${SERVER_PORT}`;
const BATCH_SIZE = Number.parseInt(process.env.BATCH_SIZE || '5', 10);
const LOAD_TIMEOUT = Number.parseInt(process.env.LOAD_TIMEOUT || '5000', 10);
const SETTLE_MS = Number.parseInt(process.env.SETTLE_MS || '10000', 10);

const DEFAULT_LIST_PATH = path.join(__dirname, 'test', 'batch-files.txt');
const LIST_PATH = process.env.FILE_LIST || process.argv[2] || DEFAULT_LIST_PATH;
const RUN_ID = `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;

if (!Number.isFinite(BATCH_SIZE) || BATCH_SIZE < 1) {
  console.log('BATCH_SIZE must be a positive integer');
  process.exit(1);
}

if (!Number.isFinite(LOAD_TIMEOUT) || LOAD_TIMEOUT < 1000) {
  console.log('LOAD_TIMEOUT must be >= 1000 ms');
  process.exit(1);
}

if (!Number.isFinite(SETTLE_MS) || SETTLE_MS < 0) {
  console.log('SETTLE_MS must be >= 0 ms');
  process.exit(1);
}

function writeStdout(line) {
  try {
    fs.writeSync(1, line + '\n');
  } catch (err) {
    process.stdout.write(line + '\n');
  }
}

function httpRequest(options) {
  return new Promise((resolve, reject) => {
    const req = http.request(options, (res) => {
      const chunks = [];
      res.on('data', (chunk) => chunks.push(chunk));
      res.on('end', () => {
        resolve({
          status: res.statusCode,
          body: Buffer.concat(chunks).toString('utf8')
        });
      });
    });
    req.on('error', reject);
    req.end();
  });
}

async function checkServer() {
  try {
    const response = await httpRequest({
      hostname: 'localhost',
      port: SERVER_PORT,
      path: '/healthcheck',
      method: 'GET'
    });
    if (response.status !== 200) {
      throw new Error(`Healthcheck failed with status ${response.status}`);
    }
  } catch (error) {
    writeStdout(`Server is not running at ${SERVER_URL}`);
    writeStdout('Start the server with: npm run server');
    process.exit(1);
  }
}

function normalizeListEntry(raw) {
  if (!raw) return null;
  let value = raw.trim();
  if (!value || value.startsWith('#')) return null;
  if ((value.startsWith('"') && value.endsWith('"')) || (value.startsWith("'") && value.endsWith("'"))) {
    value = value.slice(1, -1).trim();
  }
  if (!value) return null;
  if (value === '~') {
    value = os.homedir();
  } else if (value.startsWith('~/')) {
    value = path.join(os.homedir(), value.slice(2));
  }
  if (!path.isAbsolute(value)) {
    value = path.resolve(process.cwd(), value);
  }
  return value;
}

function loadFileList(listPath) {
  if (!fs.existsSync(listPath)) {
    writeStdout(`File list not found: ${listPath}`);
    writeStdout('Provide a list file path as the first argument or set FILE_LIST.');
    process.exit(1);
  }
  const lines = fs.readFileSync(listPath, 'utf8').split(/\r?\n/);
  const files = lines
    .map(normalizeListEntry)
    .filter((entry) => entry);
  if (files.length === 0) {
    writeStdout(`File list is empty: ${listPath}`);
    process.exit(1);
  }
  return files;
}

function validateFilePaths(files) {
  const missing = [];
  for (const filePath of files) {
    if (!fs.existsSync(filePath) || !fs.statSync(filePath).isFile()) {
      missing.push(filePath);
    }
  }
  if (missing.length > 0) {
    writeStdout('The following files do not exist:');
    missing.forEach((filePath) => writeStdout(`- ${filePath}`));
    process.exit(1);
  }
}

function buildBaseEntry(filePath, fileIndex) {
  return {
    ts: new Date().toISOString(),
    runId: RUN_ID,
    fileIndex,
    filePath
  };
}

let hasFailed = false;
let browser = null;

function fail(reason, entry) {
  if (hasFailed) return;
  hasFailed = true;

  const payload = {
    ...entry,
    reason,
    fatal: true
  };
  writeStdout('[error] browser log error detected');
  writeStdout(JSON.stringify(payload));

  process.exit(1);
}

function attachPageLogging(page, filePath, fileIndex) {
  page.on('console', (msg) => {
    const type = msg.type();
    if (type !== 'error' && type !== 'assert') {
      return;
    }
    const location = msg.location();
    const entry = {
      ...buildBaseEntry(filePath, fileIndex),
      event: 'console',
      level: type,
      text: msg.text(),
      url: location.url || null,
      line: Number.isFinite(location.lineNumber) ? location.lineNumber : null,
      column: Number.isFinite(location.columnNumber) ? location.columnNumber : null
    };
    fail('console error', entry);
  });

  page.on('pageerror', (error) => {
    const entry = {
      ...buildBaseEntry(filePath, fileIndex),
      event: 'pageerror',
      level: 'error',
      text: error.message,
      stack: error.stack || null
    };
    fail('pageerror', entry);
  });

  page.on('crash', () => {
    const entry = {
      ...buildBaseEntry(filePath, fileIndex),
      event: 'crash',
      level: 'error',
      text: 'Page crashed'
    };
    fail('page crashed', entry);
  });
}

async function openDocument(filePath, fileIndex) {
  if (hasFailed) return;

  const context = await browser.newContext({
    viewport: { width: 1280, height: 720 }
  });
  const page = await context.newPage();

  attachPageLogging(page, filePath, fileIndex);

  const url = `${SERVER_URL}/open?filepath=${encodeURIComponent(filePath)}`;

  try {
    await page.goto(url, { waitUntil: 'load', timeout: LOAD_TIMEOUT });
    await page.waitForTimeout(SETTLE_MS);
  } catch (error) {
    if (!hasFailed) {
      const entry = {
        ...buildBaseEntry(filePath, fileIndex),
        event: 'navigation_error',
        level: 'error',
        text: error.message,
        url
      };
      fail('navigation error', entry);
    }
  } finally {
    try {
      await context.close();
    } catch (err) {
      if (!hasFailed) {
        writeStdout(`Context close warning: ${err.message}`);
      }
    }
  }
}

async function run() {
  await checkServer();

  const files = loadFileList(LIST_PATH);
  validateFilePaths(files);

  writeStdout(`Batch run ${RUN_ID} starting (${files.length} files, batch size ${BATCH_SIZE})`);
  writeStdout(`Server: ${SERVER_URL}`);
  writeStdout(`List: ${LIST_PATH}`);

  browser = await chromium.launch({ headless: true });

  for (let i = 0; i < files.length; i += BATCH_SIZE) {
    if (hasFailed) return;
    const batch = files.slice(i, i + BATCH_SIZE);
    await Promise.all(batch.map((filePath, idx) => openDocument(filePath, i + idx)));
  }

  await browser.close();
  writeStdout('Batch console log check completed with no errors.');
  process.exit(0);
}

run().catch((error) => {
  writeStdout(`Unhandled error: ${error.message}`);
  process.exit(1);
});
