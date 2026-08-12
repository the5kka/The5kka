const {chromium}=require('playwright');
const path=require('path');
(async()=>{
  const browser=await chromium.launch({headless:true,executablePath:'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe'});
  const page=await browser.newPage({viewport:{width:1800,height:1200},deviceScaleFactor:1});
  const source='file:///'+path.join(__dirname,'contact_sheet.html').replace(/\\/g,'/');
  for(const item of [{q:'?page=1',name:'전체_비교_01.png',height:1200},{q:'?page=2',name:'전체_비교_02.png',height:1200},{q:'?mode=picks',name:'추천_8종.png',height:760}]){
    await page.setViewportSize({width:1800,height:item.height});
    await page.goto(source+item.q,{waitUntil:'load'});
    await page.evaluate(()=>document.fonts.ready);
    await page.screenshot({path:path.join(__dirname,item.name)});
  }
  await browser.close();
})().catch(e=>{console.error(e);process.exit(1)});
