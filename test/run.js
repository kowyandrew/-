const { chromium } = require('playwright');
const fs = require('fs');
const HTML = fs.readFileSync('../index.html','utf8');
const STUB = fs.readFileSync('./stub.js','utf8');
fs.writeFileSync('/tmp/board.html',
  '<!doctype html><html><head><meta charset="utf-8">'+
  '<style>*{box-sizing:border-box}body{margin:0;font:14px system-ui}img{max-width:100%}[hidden]{display:none!important}</style>'+
  '</head><body>'+HTML+'</body></html>');
const URL='file:///tmp/board.html';
const step=m=>console.log(m);

(async()=>{
  const b=await chromium.launch({executablePath:'/opt/pw-browsers/chromium-1194/chrome-linux/chrome'});
  const errors=[];
  const ctx=await b.newContext({viewport:{width:1024,height:1366},deviceScaleFactor:2});
  const p=await ctx.newPage();
  p.on('pageerror',e=>errors.push('JS: '+e.message));
  p.on('console',m=>{const t=m.text(); if(m.type()==='error'&&!/ERR_CONNECTION|font/i.test(t))errors.push('console: '+t);});
  await p.addInitScript(STUB);
  await p.goto(URL,{waitUntil:'domcontentloaded'});
  await p.waitForTimeout(300);

  const be=async who=>{ await p.click('#whoBtn'); await p.waitForSelector('.sheet');
    await p.click(`.chip[data-name="${who}"]`); await p.waitForTimeout(300); };

  await be('専務');
  step('① 本人設定: '+await p.textContent('#whoName'));

  for(const[t,k,ph,h,a]of[['新宿','建売','2','拓也','外構の図面を今週中に出す'],
                          ['大宮2号地','建売','4','山口','基礎の配筋検査の日程を押さえる'],
                          ['桶川の土地','土地','0','渡邉','売主と決済日をつめる']]){
    await p.click('#newBtn'); await p.waitForSelector('#f_title');
    await p.fill('#f_title',t); await p.selectOption('#f_kind',k); await p.selectOption('#f_phase',ph);
    await p.click(`#ncHolder .chip[data-name="${h}"]`); await p.fill('#f_ask',a);
    await p.click('#f_ok'); await p.waitForTimeout(250);
  }
  step('② 案件数: '+await p.locator('.case').count());

  // 専務が渡邉に質問
  await p.click('.case >> nth=0 >> .case-head'); await p.waitForTimeout(150);
  await p.click('[data-act="ask"]'); await p.waitForSelector('#q_t');
  await p.click('#qTo .chip[data-name="渡邉"]');
  await p.fill('#q_t','この土地の決済日はもう決まってる？'); await p.click('#q_ok');
  await p.waitForTimeout(350);
  step('③ 質問バンド:'+await p.locator('.qband').count()+' 未回答フィルタ:'+await p.textContent('[data-filter="q"] .n')
      +' / 注記:'+(await p.locator('.pending-note').count()?(await p.textContent('.pending-note')).trim():'なし'));
  step('④ 未回答案件が先頭か: '+(await p.locator('.case').first().locator('.title').textContent()));
  await p.screenshot({path:'01-senmu-ipad.png',fullPage:true});

  // 渡邉に切り替えて返答
  await be('渡邉');
  const q=await p.locator('.mine.q').count();
  step('⑤ 渡邉「あなたへの質問」: '+q+' 件 → '+(q?(await p.textContent('.mine.q .minerow .a')).trim():'なし'));
  await p.screenshot({path:'02-watanabe.png',fullPage:true});
  await p.click('.mine.q .minerow'); await p.waitForTimeout(250);
  await p.click('.qband [data-answer]'); await p.waitForSelector('#a_t');
  await p.fill('#a_t','9/12で押さえました。司法書士も手配済みです。');
  await p.click('#a_ok'); await p.waitForTimeout(350);
  step('⑥ 返答後の質問バンド: '+await p.locator('.qband').count()+'（0が正）');

  // 専務に戻ると返答が届いている
  await be('専務');
  const back=await p.locator('.mine.a').count();
  step('⑦ 専務「返答がありました」: '+back+' → '+(back?(await p.textContent('.mine.a .minerow .a')).trim():'なし'));
  await p.screenshot({path:'03-reply-back.png',fullPage:true});
  await p.click('.mine.a .minerow'); await p.waitForTimeout(300);
  step('⑧ 開いた後に返答ブロックが消えるか: '+(await p.locator('.mine.a').count()===0?'消えた':'残っている'));
  step('⑨ 履歴の先頭: '+(await p.locator('.log li').first().innerText()).replace(/\n/g,' / '));

  // Chatwork 文面
  await p.click('.tcard[data-holder="山口"]'); await p.waitForTimeout(250);
  await p.click('[data-copy]'); await p.waitForSelector('#sh_t');
  step('⑩ Chatwork 文面 ----------\n'+await p.inputValue('#sh_t')+'\n----------');
  await p.keyboard.press('Escape'); await p.waitForTimeout(150);

  // 名簿
  await p.click('#rosterBtn'); await p.waitForSelector('.roster');
  await p.fill('[data-role-for="酒井"]','仕入れ・銀行');
  await p.click('#ros_ok'); await p.waitForTimeout(300);
  await p.click('.tcard >> nth=0'); await p.waitForTimeout(200);
  step('⑪ 名簿保存後の役割: '+JSON.stringify((await p.evaluate(()=>window.__dump()))['meta/members'].roles));

  // iPhone
  const p2=await(await b.newContext({viewport:{width:390,height:844},deviceScaleFactor:2})).newPage();
  p2.on('pageerror',e=>errors.push('iphone JS: '+e.message));
  await p2.addInitScript(STUB); await p2.goto(URL,{waitUntil:'domcontentloaded'}); await p2.waitForTimeout(300);
  await p2.click('#whoBtn'); await p2.waitForSelector('.sheet'); await p2.click('.chip[data-name="山口"]');
  await p2.waitForTimeout(250);
  await p2.screenshot({path:'04-iphone-empty.png',fullPage:true});

  // ダーク
  const pd=await(await b.newContext({viewport:{width:1024,height:900},colorScheme:'dark',deviceScaleFactor:2})).newPage();
  pd.on('pageerror',e=>errors.push('dark JS: '+e.message));
  await pd.addInitScript(STUB); await pd.goto(URL,{waitUntil:'domcontentloaded'}); await pd.waitForTimeout(400);
  step('⑫ dark body bg: '+await pd.evaluate(()=>getComputedStyle(document.body).backgroundColor));
  await pd.screenshot({path:'05-dark.png'});

  for(const[n,pg]of[['ipad',p],['iphone',p2]]){
    step('⑬ '+n+' 横あふれ: '+await pg.evaluate(()=>document.documentElement.scrollWidth-document.documentElement.clientWidth)+'px');
  }
  step(errors.length?'❌ ERRORS:\n'+errors.join('\n'):'✅ JSエラーなし');
  await b.close();
})().catch(e=>{console.error('FAILED',e.message);process.exit(1);});
