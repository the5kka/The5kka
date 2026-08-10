const {chromium}=require('playwright');
const fs=require('fs');
const path=require('path');
(async()=>{
  const browser=await chromium.launch({headless:true,executablePath:'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe'});
  const page=await browser.newPage({viewport:{width:1600,height:900},deviceScaleFactor:1});
  const source='file:///'+path.join(__dirname,'research_lab.html').replace(/\\/g,'/');
  const report=[];
  for(let id=25;id<=60;id++){
    const pageErrors=[];
    const handler=e=>pageErrors.push(e.message);
    page.on('pageerror',handler);
    await page.goto(`${source}?d=${id}`,{waitUntil:'load'});
    await page.waitForSelector('#app > *',{timeout:10000});
    await page.evaluate(()=>document.fonts.ready);
    const issues=await page.evaluate(()=>{
      const found=[],app=document.querySelector('#app'),ar=app.getBoundingClientRect(),tol=2;
      const visible=el=>{const s=getComputedStyle(el),r=el.getBoundingClientRect();return s.display!=='none'&&s.visibility!=='hidden'&&r.width>0&&r.height>0};
      document.querySelectorAll('.btn,.control,.role-card,.tile,.task-row,.safety-item,.fluent-card,.command-box,.command-result,.carbon-row,.ix-tile,.map-line,.fiori-tile,.space-app,.us-option,.high-choice,.scan-box,.journey-step,.ticket-choice,.library-row,.pos-item,.atm-option,.key,.passport-card,.passport-action,.roster-row,.class-card,.credential,.badge-reader,.qr-frame,.favorite,.time-card,.map-node,.andon-cell,.shift-tab,.event,.paper,.dock-panel,.dock-action,.oneq-answer').forEach((el,n)=>{
        if(!visible(el))return;const r=el.getBoundingClientRect();
        if(r.left<ar.left-tol||r.top<ar.top-tol||r.right>ar.right+tol||r.bottom>ar.bottom+tol)found.push(`outside:${el.className}:${n}`);
      });
      document.querySelectorAll('.paper-stage,.print-stage,.compare-paper').forEach((stage,n)=>{const p=stage.querySelector('.paper');if(!p||!visible(p))return;const s=stage.getBoundingClientRect(),r=p.getBoundingClientRect();if(r.left<s.left-tol||r.top<s.top-tol||r.right>s.right+tol||r.bottom>s.bottom+tol)found.push(`paper-outside:${n}`);const ratio=r.width/r.height;if(Math.abs(ratio-210/297)>.02)found.push(`paper-ratio:${ratio.toFixed(3)}`)});
      document.querySelectorAll('.btn,.control,.role-pill,.atm-option,.dock-action').forEach((el,n)=>{if(visible(el)&&(el.scrollWidth>el.clientWidth+3||el.scrollHeight>el.clientHeight+3))found.push(`control-overflow:${n}`)});
      if(document.body.innerText.includes('${'))found.push('template-visible');
      if(document.querySelectorAll('svg.lucide').length<1)found.push('icons-missing');
      if(app.scrollWidth>1602||app.scrollHeight>902)found.push(`app-scroll:${app.scrollWidth}x${app.scrollHeight}`);
      return [...new Set(found)];
    });
    page.off('pageerror',handler);
    report.push({design:id,issues:[...pageErrors.map(x=>`page-error:${x}`),...issues]});
  }
  await browser.close();
  fs.writeFileSync(path.join(__dirname,'qa_report.json'),JSON.stringify(report,null,2));
  const failed=report.filter(x=>x.issues.length);
  console.log(JSON.stringify({checked:36,passed:36-failed.length,failed:failed.length,details:failed},null,2));
  if(failed.length)process.exitCode=2;
})().catch(e=>{console.error(e);process.exit(1)});
