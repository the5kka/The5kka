const {chromium}=require('playwright');
const path=require('path');
(async()=>{
  const requested=process.argv.slice(2).flatMap(x=>x.split(',')).map(Number).filter(n=>n>=25&&n<=60);
  const ids=requested.length?requested:Array.from({length:36},(_,i)=>i+25);
  const browser=await chromium.launch({headless:true,executablePath:'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe'});
  const page=await browser.newPage({viewport:{width:1600,height:900},deviceScaleFactor:1});
  const source='file:///'+path.join(__dirname,'research_lab.html').replace(/\\/g,'/');
  for(const id of ids){
    await page.goto(`${source}?d=${id}`,{waitUntil:'load'});
    await page.waitForSelector('#app > *',{timeout:10000});
    await page.evaluate(()=>document.fonts.ready);
    await page.screenshot({path:path.join(__dirname,'screens',`design_${id}.png`)});
    console.log(`rendered ${id}`);
  }
  await browser.close();
})().catch(e=>{console.error(e);process.exit(1)});
