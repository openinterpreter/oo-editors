import { chromium } from 'playwright';
import { spawnSync } from 'node:child_process';
import fs from 'node:fs';
import path from 'node:path';

const file = path.resolve(process.argv[2] || '.github/assets/simple.xlsx');
const port = process.env.OO_PORT || '38123';
const url = `http://localhost:${port}/open?filepath=${encodeURIComponent(file)}`;

function findChromeExecutable() {
  const pathCandidates = [
    'google-chrome-stable',
    'google-chrome',
    'chromium',
    'chromium-browser',
    'chrome',
    'brave-browser',
    'microsoft-edge'
  ];
  const candidates = [
    process.env.CHROME_BIN,
    process.env.GOOGLE_CHROME_SHIM,
    ...pathCandidates.map(command => {
      const result = spawnSync('command', ['-v', command], { shell: true, encoding: 'utf8' });
      return result.status === 0 ? result.stdout.trim() : '';
    }),
    '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
    '/Applications/Google Chrome Canary.app/Contents/MacOS/Google Chrome Canary',
    '/Applications/Chromium.app/Contents/MacOS/Chromium',
    path.join(process.env.HOME || '', 'Applications/Home Manager Apps/Google Chrome.app/Contents/MacOS/Google Chrome'),
    '/Applications/Arc.app/Contents/MacOS/Arc',
    '/usr/bin/google-chrome',
    '/usr/bin/google-chrome-stable',
    '/usr/bin/chromium',
    '/usr/bin/chromium-browser',
    'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',
    'C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe'
  ].filter(Boolean);

  return candidates.find(candidate => fs.existsSync(candidate));
}

const chromeExecutable = findChromeExecutable();
if (!chromeExecutable) {
  throw new Error('No global Chrome/Chromium executable found. Set CHROME_BIN=/path/to/chrome and rerun.');
}

const launchOptions = {
  headless: false,
  executablePath: chromeExecutable,
  args: [
    '--disable-crash-reporter',
    '--disable-crashpad'
  ]
};

const browser = await chromium.launch(launchOptions);
const page = await browser.newPage({ viewport: { width: 1400, height: 900 } });

await page.exposeFunction('printOOEvent', event => {
  console.log(JSON.stringify(event, null, 2));
});

await page.setContent(`
  <!doctype html>
  <meta charset="utf-8">
  <style>
    html, body, iframe { margin: 0; width: 100%; height: 100%; border: 0; }
  </style>
  <script>
    window.addEventListener('message', event => {
      window.printOOEvent({
        origin: event.origin,
        data: event.data
      });
    });
  </script>
  <iframe src="${url}"></iframe>
`);

console.log(`Using Chrome: ${chromeExecutable || 'Playwright channel chrome'}`);
console.log(`Listening for oo-editors events from: ${file}`);
console.log('Select cells/text/images in the opened editor. Ctrl+C to stop.');
await new Promise(() => {});
