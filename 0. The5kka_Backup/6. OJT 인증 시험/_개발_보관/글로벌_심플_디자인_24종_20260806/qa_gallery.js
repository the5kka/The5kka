const {chromium}=require('playwright');
const fs=require('fs');
const path=require('path');

(async()=>{
  const browser=await chromium.launch({headless:true,executablePath:'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe'});
  const galleryPath=process.argv[2]?path.resolve(process.argv[2]):path.join(__dirname,'00_OJT_글로벌_심플_디자인_24종.html');
  const source='file:///'+galleryPath.replace(/\\/g,'/');
  const results=[];

  const desktop=await browser.newPage({viewport:{width:1440,height:900},deviceScaleFactor:1});
  await desktop.goto(source,{waitUntil:'load'});
  await desktop.waitForSelector('.card');
  const desktopCheck=await desktop.evaluate(()=>({
    cards:document.querySelectorAll('.card').length,
    broken:[...document.querySelectorAll('.thumb img')].filter(x=>!x.complete||x.naturalWidth!==1600||x.naturalHeight!==900).length,
    horizontalOverflow:document.documentElement.scrollWidth>window.innerWidth+2
  }));
  await desktop.locator('.card').first().click();
  await desktop.waitForSelector('.modal.open');
  const modalBefore=await desktop.locator('#modalTitle').innerText();
  await desktop.locator('#next').click();
  const modalAfter=await desktop.locator('#modalTitle').innerText();
  const modalImage=await desktop.locator('#modalImage').evaluate(x=>({width:x.naturalWidth,height:x.naturalHeight,complete:x.complete}));
  await desktop.locator('#close').click();
  await desktop.screenshot({path:path.join(__dirname,'갤러리_첫화면.png')});
  results.push({name:'desktop',...desktopCheck,modalNavigation:modalBefore!==modalAfter,modalImage});

  const mobile=await browser.newPage({viewport:{width:390,height:844},deviceScaleFactor:1});
  await mobile.goto(source,{waitUntil:'load'});
  await mobile.waitForSelector('.card');
  const mobileCheck=await mobile.evaluate(()=>({
    cards:document.querySelectorAll('.card').length,
    broken:[...document.querySelectorAll('.thumb img')].filter(x=>!x.complete||x.naturalWidth!==1600||x.naturalHeight!==900).length,
    horizontalOverflow:document.documentElement.scrollWidth>window.innerWidth+2
  }));
  await mobile.screenshot({path:path.join(__dirname,'갤러리_모바일.png')});
  results.push({name:'mobile',...mobileCheck});

  await browser.close();
  const failures=[];
  for(const r of results){if(r.cards!==24)failures.push(`${r.name}: card count ${r.cards}`);if(r.broken)failures.push(`${r.name}: broken images ${r.broken}`);if(r.horizontalOverflow)failures.push(`${r.name}: horizontal overflow`)}
  if(!results[0].modalNavigation)failures.push('desktop: modal navigation');
  if(!results[0].modalImage.complete||results[0].modalImage.width!==1600||results[0].modalImage.height!==900)failures.push('desktop: modal image');
  fs.writeFileSync(path.join(__dirname,'qa_gallery.json'),JSON.stringify({results,failures},null,2));
  console.log(JSON.stringify({passed:failures.length===0,results,failures},null,2));
  if(failures.length)process.exitCode=2;
})().catch(e=>{console.error(e);process.exit(1)});
