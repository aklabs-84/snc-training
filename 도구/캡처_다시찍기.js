/* 앱 화면 캡처를 다시 찍어서 수업가이드/images/ 에 저장해요.
   준비: npm i playwright (처음 한 번)  →  실행: node 도구/캡처_다시찍기.js
   앱을 고쳤다면 먼저 `cd 시안/앱3종 && python3 build.py` 로 앱 파일을 다시 만드세요. */
const {chromium}=require('playwright');
const path=require('path');
const ROOT=path.resolve(__dirname,'..');
const OUT=path.join(ROOT,'수업가이드/images/')+'/';
const S=path.join(ROOT,'Gemini시험/')+'/';
const A='file://'+path.join(ROOT,'시안/앱3종')+'/';
(async()=>{const b=await chromium.launch();
const ctx=await b.newContext({viewport:{width:1100,height:820},colorScheme:'light'});const p=await ctx.newPage();
const errs=[];p.on('pageerror',e=>errs.push(e.message));
const fs=require('fs');
const CSV={name:'시험용_정리된_매출.csv',mimeType:'text/csv',buffer:fs.readFileSync(S+'시험용_정리된_매출.csv')}; /* 한글 파일 이름은 경로 대신 내용으로 넘겨야 안정적이에요 */
const shot=async n=>{await p.waitForTimeout(250);await p.screenshot({path:OUT+n+'.png'});console.log('저장',n)};
const top=async()=>{await p.evaluate(()=>window.scrollTo(0,0))};
const to=async sel=>{await p.evaluate(s=>{const e=document.querySelector(s);const r=e.getBoundingClientRect();window.scrollTo(0,window.scrollY+r.top-70)},sel)};
const region=async(a,z,n,pad=16)=>{const r=await p.evaluate(([a,z])=>{const x=document.querySelector(a).getBoundingClientRect(),y=document.querySelector(z).getBoundingClientRect();return{top:x.top+scrollY,bottom:y.bottom+scrollY,left:0,width:document.documentElement.clientWidth}},[a,z]);
  await p.screenshot({path:OUT+n+'.png',fullPage:true,clip:{x:0,y:Math.max(0,r.top-pad),width:r.width,height:Math.min(r.bottom-r.top+2*pad,1400)}});console.log('저장',n)};
// ===== 앱①
await p.goto(A+'1_정리집계앱.html');await p.click('[data-s=sales]');await top();await shot('앱1_01_연습자료_올린직후');
await to('#tabs');await shot('앱1_02_이상한곳_목록');
await p.click('[data-k=all]');await to('#tabs');await shot('앱1_03_추천후_이상한곳3');
await p.click('[data-a=step][data-n="3"]');await p.waitForTimeout(300);await region('#sumline','#aout','앱1_04_합계_지우기전');
await p.click('#sumline [data-a=step]');await p.click('[data-a=tool][data-k=dup]');await p.click('[data-a=step][data-n="3"]');await p.waitForTimeout(300);await region('#sumline','#aout','앱1_05_합계_지운뒤');
await p.goto(A+'1_정리집계앱.html');await p.fill('#paste','지점\t상품\t수량\t금액\n서울\t사과\t3개\t12,000원\n서울시\t사과\t2\t8000\n서울 본점\t배\t5\t15000\n부산\t배\t5\t15000\n부산시\t포도\t1\t5000\n서울\t포도\t4\t9000');await p.waitForSelector('#v2:not([hidden])');
await p.click('[data-a=gosim]');await to('#tabs');await shot('앱1_06_값목록_합치기');
await p.goto(A+'1_정리집계앱.html');await p.click('[data-s=sales]');await p.click('[data-k=all]');await p.click('[data-a=tool][data-k=dup]');await p.click('[data-a=col][data-c="1"]');await p.click('[data-a=tab][data-k=find]');await p.fill('#fnd','Sakura Parts');await p.fill('#rep','사쿠라파츠');await to('#tabs');await shot('앱1_07_바꾸기탭');
await p.click('[data-a=step][data-n="3"]');await p.selectOption('#a-split','2');await p.waitForTimeout(300);await region('#actl','#achart','앱1_08_합계_구분별');
await p.fill('#v3 .rname','3분기 매출 정리');await p.click('[data-a=rsave]');await p.waitForTimeout(300);await region('#rcres','#rcres','앱1_09_정리방법_저장');
// ===== 앱②
await p.goto(A+'2_분석시각화앱.html');await p.setInputFiles('#file',CSV);await p.waitForSelector('#v2:not([hidden])');await top();await shot('앱2_01_분석보고서_위쪽');
await p.evaluate(()=>{const h=[...document.querySelectorAll('#report h2')].find(x=>x.textContent.includes('변화는 어디에서'));window.scrollTo(0,h.getBoundingClientRect().top+scrollY-30)});await shot('앱2_02_변화는어디에서');
await p.click('[data-a=step][data-n="3"]');await p.click('input[name=pk][value=read] + span');await top();await shot('앱2_03_가리기와요청문');
await p.fill('#ans','거래처A는 9월에 12,415,000원으로 늘었고, 거래처B는 364,000원으로 크게 줄었어요.');await to('#ans');await shot('앱2_04_답붙여넣기_이름되돌림');
await p.fill('#ans','거래처A가 20,441,000원으로 1위예요. 8월은 +11.8%, 그런데 합계는 58,000,000원으로 보여요.');await to('#ans');await shot('앱2_05_숫자검사');
// ===== 앱③
await p.goto(A+'3_보고서앱.html');await p.setInputFiles('#file',CSV);await p.waitForSelector('#v2:not([hidden])');
await p.fill('#title','3분기 매출 보고');await p.fill('#period','2026년 3분기');await p.fill('#memo','수출이 늘어난 점을 강조하고 싶어요');await top();await shot('앱3_01_설정');
await p.click('[data-a=step][data-n="3"]');await top();await shot('앱3_02_요청문_내용먼저');
await p.click('[data-a=step][data-n="4"]');await p.fill('#ans','<!doctype html><html><body><h1>3분기 매출 보고</h1><p>Sakura Parts가 20,441,000원(35.5%)으로 1위입니다.</p></body></html>'.replace('Sakura Parts','거래처A'));await top();await shot('앱3_03_다듬기_웹페이지');
console.log('오류',errs);await b.close()})()
