const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');

(async () => {
  const root = __dirname;
  const screens = path.join(root, 'screens');
  fs.mkdirSync(screens, { recursive: true });
  const requested = process.argv.slice(2).flatMap(v => v.split(',')).map(Number).filter(n => n >= 1 && n <= 120);
  const designs = requested.length ? requested : Array.from({ length: 120 }, (_, i) => i + 1);
  const roleNames = ['general', 'electrical', 'newcomer', 'foreigner'];
  const browser = await chromium.launch({
    headless: true,
    executablePath: 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe'
  });
  const page = await browser.newPage({ viewport: { width: 1600, height: 900 }, deviceScaleFactor: 1 });
  page.on('pageerror', error => console.error(`PAGE ERROR: ${error.message}`));
  const source = 'file:///' + path.join(root, 'design_lab.html').replace(/\\/g, '/');
  for (const i of designs) {
    await page.goto(`${source}?d=${i}`, { waitUntil: 'load' });
    await page.waitForSelector('#app > *', { timeout: 10000 });
    await page.evaluate(() => document.fonts.ready);
    await page.waitForTimeout(120);
    if (process.env.DEBUG_LAYOUT === '1') {
      const boxes = await page.evaluate(() => [...document.querySelectorAll('.doc-stage')].map(stage => {
        const paper = stage.querySelector('.doc-page');
        const s = stage.getBoundingClientRect();
        const p = paper ? paper.getBoundingClientRect() : null;
        return { stage: { x:s.x, y:s.y, width:s.width, height:s.height }, paper: p && { x:p.x, y:p.y, width:p.width, height:p.height } };
      }));
      console.log(JSON.stringify({ design: i, boxes }, null, 2));
    }
    const role = roleNames[(i - 1) % 4];
    const file = `design_${String(i).padStart(3, '0')}_${role}.png`;
    await page.locator('#app').screenshot({ path: path.join(screens, file) });
    process.stdout.write(`rendered ${i} -> ${file}\n`);
  }
  await browser.close();
})().catch(error => {
  console.error(error);
  process.exit(1);
});
