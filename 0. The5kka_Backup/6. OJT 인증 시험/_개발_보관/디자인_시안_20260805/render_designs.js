const { chromium } = require('playwright');
const path = require('path');

(async () => {
  const root = __dirname;
  const browser = await chromium.launch({
    headless: true,
    executablePath: 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe'
  });
  const page = await browser.newPage({ viewport: { width: 1664, height: 964 }, deviceScaleFactor: 1 });
  await page.goto('file:///' + path.join(root, '디자인_시안_비교.html').replace(/\\/g, '/'), { waitUntil: 'load' });
  await page.waitForFunction(() => document.querySelectorAll('svg.lucide').length > 20);
  for (let i = 1; i <= 6; i += 1) {
    const target = page.locator(`#concept-${i}`);
    await target.screenshot({ path: path.join(root, `시안_${String(i).padStart(2, '0')}.png`) });
  }
  await browser.close();
})().catch((error) => {
  console.error(error);
  process.exit(1);
});
