const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');

(async () => {
  const root = __dirname;
  const screens = path.join(root, 'screens');
  fs.mkdirSync(screens, { recursive: true });
  const browser = await chromium.launch({
    headless: true,
    executablePath: 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe'
  });
  const page = await browser.newPage({ viewport: { width: 1480, height: 850 }, deviceScaleFactor: 1 });
  const source = 'file:///' + path.join(root, 'design_60.html').replace(/\\/g, '/');
  for (let i = 1; i <= 60; i += 1) {
    await page.goto(`${source}?i=${i}`, { waitUntil: 'load' });
    await page.waitForFunction(() => document.querySelectorAll('svg.lucide').length >= 12);
    await page.locator('#screen').screenshot({ path: path.join(screens, `design_${String(i).padStart(2, '0')}.png`) });
  }
  await browser.close();
})().catch((error) => {
  console.error(error);
  process.exit(1);
});
