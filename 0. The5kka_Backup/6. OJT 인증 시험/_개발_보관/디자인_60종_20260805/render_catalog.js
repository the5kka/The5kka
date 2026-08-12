const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');

(async () => {
  const root = __dirname;
  const out = path.join(root, 'comparison');
  fs.mkdirSync(out, { recursive: true });
  const browser = await chromium.launch({ headless: true, executablePath: 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe' });
  const page = await browser.newPage({ viewport: { width: 1800, height: 1400 }, deviceScaleFactor: 1 });
  const source = 'file:///' + path.join(root, 'gallery_60.html').replace(/\\/g, '/');
  for (let sheet = 1; sheet <= 5; sheet += 1) {
    await page.goto(`${source}?view=contact&page=${sheet}`, { waitUntil: 'load' });
    await page.waitForFunction(() => [...document.querySelectorAll('.board-only .contact-card img')].every(img => img.complete));
    await page.screenshot({ path: path.join(out, `screens_${sheet}.png`), fullPage: true });
  }
  for (const kind of ['palette', 'buttons']) {
    for (let sheet = 1; sheet <= 2; sheet += 1) {
      await page.goto(`${source}?view=${kind}&page=${sheet}`, { waitUntil: 'load' });
      if (kind === 'buttons') {
        await page.waitForFunction(() => document.querySelectorAll('.board-only .sample-button svg.lucide').length === 30);
      } else {
        await page.waitForFunction(() => document.querySelectorAll('.board-only .palette-card').length === 30);
      }
      await page.screenshot({ path: path.join(out, `${kind}_${sheet}.png`), fullPage: true });
    }
  }
  await browser.close();
})().catch(error => { console.error(error); process.exit(1); });
