const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');

(async () => {
  const root = __dirname;
  const compare = path.join(root, '비교표_12개씩');
  const recommended = path.join(root, '추천_24종');
  fs.mkdirSync(compare, { recursive: true });
  fs.mkdirSync(recommended, { recursive: true });
  const browser = await chromium.launch({ headless: true, executablePath: 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe' });
  const page = await browser.newPage({ viewport: { width: 1920, height: 1080 }, deviceScaleFactor: 1 });
  const contact = 'file:///' + path.join(root, 'contact_sheet.html').replace(/\\/g, '/');
  const waitImages = () => page.waitForFunction(() => [...document.images].every(img => img.complete && img.naturalWidth > 0));
  for (let n = 1; n <= 10; n += 1) {
    await page.goto(`${contact}?mode=all&page=${n}`, { waitUntil: 'load' });
    await waitImages();
    const first = (n - 1) * 12 + 1;
    const last = n * 12;
    await page.locator('.sheet').screenshot({ path: path.join(compare, `비교_${String(n).padStart(2,'0')}_${String(first).padStart(3,'0')}-${String(last).padStart(3,'0')}.png`) });
  }
  for (let n = 1; n <= 2; n += 1) {
    await page.goto(`${contact}?mode=recommended&page=${n}`, { waitUntil: 'load' });
    await waitImages();
    await page.locator('.sheet').screenshot({ path: path.join(recommended, `추천_24종_${n}.png`) });
  }
  await page.setViewportSize({ width: 1600, height: 1000 });
  const gallery = 'file:///' + path.join(root, '00_OJT_전체화면_디자인_120종.html').replace(/\\/g, '/');
  await page.goto(gallery, { waitUntil: 'load' });
  await page.waitForFunction(() => [...document.images].slice(0, 9).every(img => img.complete && img.naturalWidth > 0));
  await page.screenshot({ path: path.join(root, '00_갤러리_첫화면.png') });
  await browser.close();
})().catch(error => { console.error(error); process.exit(1); });
