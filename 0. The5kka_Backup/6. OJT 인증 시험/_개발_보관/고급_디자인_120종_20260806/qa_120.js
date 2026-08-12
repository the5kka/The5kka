const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');

(async () => {
  const root = __dirname;
  const browser = await chromium.launch({
    headless: true,
    executablePath: 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe'
  });
  const page = await browser.newPage({ viewport: { width: 1600, height: 900 }, deviceScaleFactor: 1 });
  const source = 'file:///' + path.join(root, 'design_lab.html').replace(/\\/g, '/');
  const report = [];
  for (let i = 1; i <= 120; i += 1) {
    const pageErrors = [];
    const handler = error => pageErrors.push(error.message);
    page.on('pageerror', handler);
    await page.goto(`${source}?d=${i}`, { waitUntil: 'load' });
    await page.waitForSelector('#app > *', { timeout: 10000 });
    await page.evaluate(() => document.fonts.ready);
    const issues = await page.evaluate(() => {
      const found = [];
      const app = document.querySelector('#app');
      const ar = app.getBoundingClientRect();
      const tol = 2;
      const visible = el => {
        const s = getComputedStyle(el);
        const r = el.getBoundingClientRect();
        return s.display !== 'none' && s.visibility !== 'hidden' && r.width > 0 && r.height > 0;
      };
      const selectors = '.btn,.panel,.topbar,.role-tabs,.doc-page,.role-card,.kiosk-tile,.condition,.field .control,.flow-node,.contrast-tile';
      document.querySelectorAll(selectors).forEach((el, n) => {
        if (!visible(el)) return;
        const r = el.getBoundingClientRect();
        if (r.left < ar.left - tol || r.top < ar.top - tol || r.right > ar.right + tol || r.bottom > ar.bottom + tol) {
          found.push(`outside:${el.className || el.tagName}:${n}`);
        }
      });
      document.querySelectorAll('.doc-stage,.paperfirst-stage').forEach((stage, n) => {
        const paper = stage.querySelector('.doc-page');
        if (!paper || !visible(paper)) return;
        const s = stage.getBoundingClientRect();
        const p = paper.getBoundingClientRect();
        if (p.left < s.left - tol || p.top < s.top - tol || p.right > s.right + tol || p.bottom > s.bottom + tol) found.push(`paper-outside:${n}`);
        const ratio = p.width / p.height;
        if (Math.abs(ratio - 210 / 297) > 0.02) found.push(`paper-ratio:${ratio.toFixed(3)}`);
      });
      document.querySelectorAll('.btn,.role-tab,.field .control,.language').forEach((el, n) => {
        if (!visible(el)) return;
        if (el.scrollWidth > el.clientWidth + 3 || el.scrollHeight > el.clientHeight + 3) found.push(`control-overflow:${n}`);
      });
      if (app.scrollWidth > 1602 || app.scrollHeight > 902) found.push(`app-scroll:${app.scrollWidth}x${app.scrollHeight}`);
      if (document.body.innerText.includes('${')) found.push('template-literal-visible');
      if (document.querySelectorAll('svg.lucide').length < 1) found.push('icons-missing');
      return [...new Set(found)];
    });
    page.off('pageerror', handler);
    report.push({ design: i, issues: [...pageErrors.map(x => `page-error:${x}`), ...issues] });
  }
  await browser.close();
  fs.writeFileSync(path.join(root, 'qa_report.json'), JSON.stringify(report, null, 2));
  const failed = report.filter(item => item.issues.length);
  console.log(JSON.stringify({ checked: report.length, passed: report.length - failed.length, failed: failed.length, details: failed }, null, 2));
  if (failed.length) process.exitCode = 2;
})().catch(error => {
  console.error(error);
  process.exit(1);
});
