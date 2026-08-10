const {chromium}=require('playwright');
const path=require('path');
(async()=>{
  const browser=await chromium.launch({headless:true,executablePath:'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe'});
  const page=await browser.newPage({viewport:{width:1800,height:1200},deviceScaleFactor:1});
  const source='file:///'+path.join(__dirname,'contact_sheet.html').replace(/\\/g,'/');
  for(let n=1;n<=4;n++){
    await page.goto(`${source}?page=${n}`,{waitUntil:'load'});
    await page.evaluate(()=>document.fonts.ready);
    await page.screenshot({path:path.join(__dirname,`전체_비교_추가_${String(n).padStart(2,'0')}.png`)});
  }
  await page.setViewportSize({width:1800,height:1260});
  const recommend='file:///'+path.join(__dirname,'recommend_sheet.html').replace(/\\/g,'/');
  await page.goto(recommend,{waitUntil:'load'});
  await page.evaluate(()=>document.fonts.ready);
  await page.screenshot({path:path.join(__dirname,'추천_16종.png')});
  await browser.close();
})().catch(e=>{console.error(e);process.exit(1)});
