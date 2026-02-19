import { chromium } from 'playwright';
import fs from 'fs';
import path from 'path';
import os from 'os';

const SERVER_PORT = Number.parseInt(process.env.SERVER_PORT || '8080', 10);
const SERVER_URL = process.env.SERVER_URL || `http://localhost:${SERVER_PORT}`;
const LOAD_TIMEOUT = Number.parseInt(process.env.LOAD_TIMEOUT || '30000', 10);
const LOGO_TIMEOUT = Number.parseInt(process.env.LOGO_TIMEOUT || '30000', 10);
const DEFAULT_LIST_PATH = path.join(import.meta.dir, '..', 'test', 'batch-files.txt');
const LIST_PATH = process.env.FILE_LIST || process.argv[2] || DEFAULT_LIST_PATH;

function normalizeListEntry(raw) {
  if (!raw) return null;
  let value = raw.trim();
  if (!value || value.startsWith('#')) return null;
  if ((value.startsWith('"') && value.endsWith('"')) || (value.startsWith("'") && value.endsWith("'"))) {
    value = value.slice(1, -1).trim();
  }
  if (!value) return null;
  if (value === '~') value = os.homedir();
  else if (value.startsWith('~/')) value = path.join(os.homedir(), value.slice(2));
  if (!path.isAbsolute(value)) value = path.resolve(process.cwd(), value);
  return value;
}

function loadFileList(listPath) {
  if (!fs.existsSync(listPath)) {
    throw new Error(`File list not found: ${listPath}`);
  }
  const lines = fs.readFileSync(listPath, 'utf8').split(/\r?\n/);
  const files = lines.map(normalizeListEntry).filter(Boolean);
  if (files.length === 0) {
    throw new Error(`File list is empty: ${listPath}`);
  }
  return files;
}

function assertFilesExist(files) {
  const missing = files.filter((filePath) => !fs.existsSync(filePath) || !fs.statSync(filePath).isFile());
  if (missing.length) {
    throw new Error(`Missing files:\n${missing.join('\n')}`);
  }
}

function getDocType(ext) {
  if (['xlsx', 'xls', 'ods', 'csv'].includes(ext)) return 'cell';
  if (['docx', 'doc', 'odt', 'txt', 'rtf', 'html'].includes(ext)) return 'word';
  if (['pptx', 'ppt', 'odp'].includes(ext)) return 'slide';
  return 'slide';
}

function buildOfflineLoaderUrl(filePath) {
  const ext = path.extname(filePath).slice(1).toLowerCase();
  const doctype = getDocType(ext);
  const convertUrl = `${SERVER_URL}/api/convert?filepath=${encodeURIComponent(filePath)}`;
  const params = new URLSearchParams({
    url: convertUrl,
    title: path.basename(filePath),
    filepath: filePath,
    filetype: ext,
    doctype,
  });
  return `${SERVER_URL}/offline-loader-proper.html?${params.toString()}`;
}

async function assertServer() {
  const response = await fetch(`${SERVER_URL}/healthcheck`).catch(() => null);
  if (!response || !response.ok) {
    throw new Error(`Server not running at ${SERVER_URL}. Start with: bun run server`);
  }
}

async function assertLogoVisible(page, filePath) {
  const url = buildOfflineLoaderUrl(filePath);
  await page.goto(url, { waitUntil: 'load', timeout: LOAD_TIMEOUT });
  await page.waitForFunction(() => {
    const iframe = document.querySelector('iframe[id*="placeholder"]') ||
      document.querySelector('iframe[id*="frameEditor"]') ||
      document.querySelector('iframe');
    if (!iframe || !iframe.contentDocument) return false;

    const doc = iframe.contentDocument;
    const logo = doc.querySelector('#oo-desktop-logo') ||
      doc.querySelector('#box-document-title .extra img[src*="header-logo_s.svg"]') ||
      doc.querySelector('#box-document-title img[src*="header-logo_s.svg"]');
    if (!logo) return false;

    const style = doc.defaultView.getComputedStyle(logo);
    const rect = logo.getBoundingClientRect();
    return style.display !== 'none' && style.visibility !== 'hidden' && rect.width > 0 && rect.height > 0;
  }, { timeout: LOGO_TIMEOUT });
}

async function run() {
  await assertServer();
  const files = loadFileList(LIST_PATH);
  assertFilesExist(files);

  const browser = await chromium.launch({ headless: true });
  const page = await browser.newPage({ viewport: { width: 1400, height: 900 } });

  try {
    for (const filePath of files) {
      await assertLogoVisible(page, filePath);
      console.log(`PASS logo visible: ${filePath}`);
    }
    console.log(`Logo header check passed for ${files.length} files`);
    await browser.close();
    process.exit(0);
  } catch (err) {
    await browser.close();
    console.error('Logo header check failed');
    console.error(err && err.message ? err.message : err);
    process.exit(1);
  }
}

run().catch((err) => {
  console.error(err && err.message ? err.message : err);
  process.exit(1);
});
