const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');

(async () => {
  const root = __dirname;
  const browser = await chromium.launch({ headless: true, executablePath: 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe' });
  const page = await browser.newPage({ viewport: { width: 1480, height: 850 }, deviceScaleFactor: 1 });
  const source = 'file:///' + path.join(root, 'design_60.html').replace(/\\/g, '/');
  const report = [];
  for (let i = 1; i <= 60; i += 1) {
    await page.goto(`${source}?i=${i}`, { waitUntil: 'load' });
    await page.waitForFunction(() => document.querySelectorAll('svg.lucide').length >= 12);
    const issues = await page.evaluate(() => {
      const found = [];
      const tolerance = 2;
      const visible = el => el && getComputedStyle(el).display !== 'none' && el.getBoundingClientRect().width > 0 && el.getBoundingClientRect().height > 0;
      const inside = (child, parent, name) => {
        if (!visible(child) || !visible(parent)) return;
        const c = child.getBoundingClientRect(), p = parent.getBoundingClientRect();
        if (c.left < p.left - tolerance || c.top < p.top - tolerance || c.right > p.right + tolerance || c.bottom > p.bottom + tolerance) found.push(`${name} outside`);
      };
      const header = document.querySelector('.header');
      document.querySelectorAll('.header > *').forEach((el, n) => inside(el, header, `header-${n}`));
      const actions = document.querySelector('.actions');
      actions.querySelectorAll('.btn,.pager,.action-buttons').forEach((el, n) => inside(el, actions, `action-${n}`));
      const cond = document.querySelector('.conditions');
      cond.querySelectorAll('.condition-body,.metrics,.total,.identity').forEach((el, n) => inside(el, cond, `condition-${n}`));
      const stage = document.querySelector('.paper-stage'), paper = document.querySelector('.paper');
      inside(paper, stage, 'paper');
      const pr = paper.getBoundingClientRect();
      const ratio = pr.width / pr.height;
      if (Math.abs(ratio - 210 / 297) > 0.025) found.push(`paper ratio ${ratio.toFixed(3)}`);
      if (pr.width < 150 || pr.height < 210) found.push(`paper too small ${Math.round(pr.width)}x${Math.round(pr.height)}`);
      document.querySelectorAll('.metric').forEach((el,n) => {
        if (el.scrollWidth > el.clientWidth + tolerance || el.scrollHeight > el.clientHeight + tolerance) found.push(`metric-${n} overflow`);
      });
      return [...new Set(found)];
    });
    report.push({ design: i, issues });
  }
  await browser.close();
  fs.writeFileSync(path.join(root, 'qa_report.json'), JSON.stringify(report, null, 2));
  const failed = report.filter(x => x.issues.length);
  console.log(JSON.stringify({ checked: report.length, failed: failed.length, details: failed }, null, 2));
  if (failed.length) process.exitCode = 2;
})().catch(error => { console.error(error); process.exit(1); });
