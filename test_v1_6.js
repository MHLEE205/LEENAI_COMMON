/* leenai-common.js v1.6 jsdom 시험
   TEST1: Graph wide scope 성공 → graphError=''
   TEST2: wide 실패(consent_required) → narrow 성공 → graphError='files:consent_required', graph='ok'
   TEST3: sp.getJson 정상 반환
   TEST4: sp.getJson 404 → null
   TEST5: sp.list @odata.nextLink 페이지 넘김 → 합산
   + 이전 v1.5 테스트 재실행 */
'use strict';
const {JSDOM}=require('jsdom');
const fs=require('fs');
const path=require('path');
const src=fs.readFileSync(path.join(__dirname,'v1','leenai-common.js'),'utf8');

function fakeIdToken(name,email){
  const b64=Buffer.from(JSON.stringify({name,preferred_username:email,email}),'utf8').toString('base64url');
  return 'hdr.'+b64+'.sig';
}
function def(w,k,v){Object.defineProperty(w,k,{value:v,writable:true,configurable:true});}

function makeCallbackDOM(pkceState,extraFetch){
  const dom=new JSDOM(
    '<!DOCTYPE html><html data-theme="dark"><head></head><body><div id="app" style="display:none"></div></body></html>',
    {url:'https://mhlee205.github.io/TEST/?debug=1&code=C&state='+(pkceState||'S'),runScripts:'dangerously',pretendToBeVisual:true}
  );
  const w=dom.window;
  def(w,'crypto',{subtle:{digest:async()=>new ArrayBuffer(32)},getRandomValues:(a)=>{for(let i=0;i<a.length;i++)a[i]=i%256;return a;}});
  const ss={_d:{'leenai_pkce_Z-05':JSON.stringify({verifier:'v',state:pkceState||'S'})},getItem(k){return this._d[k]||null;},setItem(k,v){this._d[k]=v;},removeItem(k){delete this._d[k];}};
  def(w,'sessionStorage',ss);
  def(w,'localStorage',{_d:{},getItem(k){return this._d[k]||null;},setItem(k,v){this._d[k]=v;},removeItem(k){delete this._d[k];}});
  def(w,'history',{replaceState:()=>{}});
  def(w,'requestAnimationFrame',(cb)=>setTimeout(cb,0));
  if(extraFetch) def(w,'fetch',extraFetch);
  return dom;
}

let passed=0,failed=0;
function ok(label,cond){
  if(cond){console.log('  ✔ '+label);passed++;}
  else{console.error('  ✖ FAIL: '+label);failed++;}
}

/* ── TEST 1: wide scope 성공 ── */
async function test1(){
  console.log('\n[TEST 1] Graph wide scope 성공 → graphError 비어있음');
  const dom=makeCallbackDOM('S1',async function(_url,opts){
    const body=opts&&opts.body||'';
    if(body.includes('authorization_code'))
      return {ok:true,json:async()=>({access_token:'DV0',expires_in:3600,refresh_token:'RT0',id_token:fakeIdToken('TEST','t@t.com')})};
    if(body.includes('refresh_token')&&body.includes('Files.Read.All'))
      return {ok:true,json:async()=>({access_token:'GW0',expires_in:3600,refresh_token:'RT1'})};
    return {ok:false,json:async()=>({error:'x'})};
  });
  dom.window.eval(src);
  const L=dom.window.LEENAI;
  await L.init({sys:'Z-05',name:'T',version:'1',date:'2026-10-08',onReady:()=>{}});
  await new Promise(r=>setTimeout(r,80));
  ok('DV ok',L.auth.status.dataverse==='ok');
  ok('Graph ok',L.auth.status.graph==='ok');
  ok('graphError empty',L.auth.status.graphError==='');
  ok('Graph token GW0',await L.auth.token('graph')==='GW0');
}

/* ── TEST 2: wide 실패 → narrow 폴백 ── */
async function test2(){
  console.log('\n[TEST 2] wide 실패(consent_required) → narrow 성공 → graphError=files:consent_required');
  const dom=makeCallbackDOM('S2',async function(_url,opts){
    const body=opts&&opts.body||'';
    if(body.includes('authorization_code'))
      return {ok:true,json:async()=>({access_token:'DV0',expires_in:3600,refresh_token:'RT0',id_token:fakeIdToken('TEST','t@t.com')})};
    if(body.includes('refresh_token')&&body.includes('Files.Read.All'))
      return {ok:false,json:async()=>({error:'consent_required'})};
    if(body.includes('refresh_token')&&body.includes('Mail.Send')&&!body.includes('Files'))
      return {ok:true,json:async()=>({access_token:'GN0',expires_in:3600,refresh_token:'RT2'})};
    return {ok:false,json:async()=>({error:'x'})};
  });
  dom.window.eval(src);
  const L=dom.window.LEENAI;
  await L.init({sys:'Z-05',name:'T',version:'1',date:'2026-10-08',onReady:()=>{}});
  await new Promise(r=>setTimeout(r,80));
  ok('Graph ok(narrow)',L.auth.status.graph==='ok');
  ok('graphError=files:consent_required',L.auth.status.graphError==='files:consent_required');
  ok('Graph token GN0',await L.auth.token('graph')==='GN0');
  ok('DV 유지',L.auth.status.dataverse==='ok');
}

/* ── TEST 3: sp.getJson 정상 반환 ── */
async function test3(){
  console.log('\n[TEST 3] sp.getJson 정상 반환');
  const lcObj={lcNo:'LC-001',country:'INDONESIA',insuranceCondition:'FOR 110 PERCENT ICC A CLAIM PAYABLE IN DESTINATION THE LC'};
  const dom=makeCallbackDOM('S3',async function(_url,opts){
    const body=opts&&opts.body||'';
    if(body&&body.includes('authorization_code'))
      return {ok:true,json:async()=>({access_token:'DV0',expires_in:3600,refresh_token:'RT0',id_token:fakeIdToken('T','t@t.com')})};
    if(body&&body.includes('refresh_token')&&body.includes('Files.Read.All'))
      return {ok:true,json:async()=>({access_token:'GW0',expires_in:3600})};
    /* sp.getJson (no body, GET) */
    if(_url&&_url.includes('drives')&&_url.includes('content'))
      return {ok:true,json:async()=>lcObj};
    return {ok:false,json:async()=>({error:'x'})};
  });
  dom.window.eval(src);
  const L=dom.window.LEENAI;
  await L.init({sys:'Z-05',name:'T',version:'1',date:'2026-10-08',onReady:()=>{}});
  await new Promise(r=>setTimeout(r,80));
  const result=await L.sp.getJson('古紙/-. AI 活用/D-01-LC_FD_CHECKER/NSG-2607-03_LC-001.json');
  ok('getJson lcNo',result&&result.lcNo==='LC-001');
  ok('getJson country',result&&result.country==='INDONESIA');
}

/* ── TEST 4: sp.getJson 404 → null ── */
async function test4(){
  console.log('\n[TEST 4] sp.getJson 404 → null');
  const dom=makeCallbackDOM('S4',async function(_url,opts){
    const body=opts&&opts.body||'';
    if(body&&body.includes('authorization_code'))
      return {ok:true,json:async()=>({access_token:'DV0',expires_in:3600,refresh_token:'RT0',id_token:fakeIdToken('T','t@t.com')})};
    if(body&&body.includes('refresh_token')&&body.includes('Files.Read.All'))
      return {ok:true,json:async()=>({access_token:'GW0',expires_in:3600})};
    if(_url&&_url.includes('drives')&&_url.includes('content'))
      return {ok:false,status:404,json:async()=>({error:{code:'itemNotFound'}})};
    return {ok:false,json:async()=>({error:'x'})};
  });
  /* fetch 응답에 status 속성 명시 필요 */
  Object.defineProperty(dom.window,'fetch',{value:async function(_url,opts){
    const body=opts&&opts.body||'';
    if(body&&body.includes('authorization_code'))
      return {ok:true,status:200,json:async()=>({access_token:'DV0',expires_in:3600,refresh_token:'RT0',id_token:fakeIdToken('T','t@t.com')})};
    if(body&&body.includes('refresh_token')&&body.includes('Files.Read.All'))
      return {ok:true,status:200,json:async()=>({access_token:'GW0',expires_in:3600})};
    if(_url&&_url.includes('drives')&&_url.includes('content'))
      return {ok:false,status:404,json:async()=>({error:{code:'itemNotFound'}})};
    return {ok:false,status:500,json:async()=>({error:'x'})};
  },writable:true,configurable:true});
  dom.window.eval(src);
  const L=dom.window.LEENAI;
  await L.init({sys:'Z-05',name:'T',version:'1',date:'2026-10-08',onReady:()=>{}});
  await new Promise(r=>setTimeout(r,80));
  const result=await L.sp.getJson('古紙/-. AI 活用/D-01-LC_FD_CHECKER/NOTEXIST.json');
  ok('getJson 404 → null',result===null);
}

/* ── TEST 5: sp.list nextLink 페이지 넘김 ── */
async function test5(){
  console.log('\n[TEST 5] sp.list @odata.nextLink 페이지 넘김 → 합산');
  let listCallCount=0;
  const dom=makeCallbackDOM('S5');
  Object.defineProperty(dom.window,'fetch',{value:async function(_url,opts){
    const body=opts&&opts.body||'';
    if(body&&body.includes('authorization_code'))
      return {ok:true,status:200,json:async()=>({access_token:'DV0',expires_in:3600,refresh_token:'RT0',id_token:fakeIdToken('T','t@t.com')})};
    if(body&&body.includes('refresh_token')&&body.includes('Files.Read.All'))
      return {ok:true,status:200,json:async()=>({access_token:'GW0',expires_in:3600})};
    if(_url&&_url.includes('children')&&!_url.includes('page2')){
      listCallCount++;
      return {ok:true,status:200,json:async()=>({
        value:[{name:'NSG-2607-03_LC-001.json',lastModifiedDateTime:'2026-08-01',size:1000},
               {name:'NSG-2607-04_LC-002.json',lastModifiedDateTime:'2026-08-02',size:1100}],
        '@odata.nextLink':'https://graph.microsoft.com/v1.0/drives/DRIVE/root:/FOLDER:/children?page2=1'
      })};
    }
    if(_url&&_url.includes('page2')){
      listCallCount++;
      return {ok:true,status:200,json:async()=>({
        value:[{name:'NSG-2607-05_LC-003.json',lastModifiedDateTime:'2026-08-03',size:1200}],
      })};
    }
    return {ok:false,status:500,json:async()=>({error:'x'})};
  },writable:true,configurable:true});
  dom.window.eval(src);
  const L=dom.window.LEENAI;
  await L.init({sys:'Z-05',name:'T',version:'1',date:'2026-10-08',onReady:()=>{}});
  await new Promise(r=>setTimeout(r,80));
  const files=await L.sp.list('古紙/-. AI 活用/D-01-LC_FD_CHECKER');
  ok('list 2페이지 호출',listCallCount===2);
  ok('list 합산 3건',files.length===3);
  ok('1번째 파일명',files[0].name==='NSG-2607-03_LC-001.json');
  ok('3번째 파일명',files[2].name==='NSG-2607-05_LC-003.json');
}

/* ── 이전 v1.5 테스트: sessionStorage debug 유지 확인 ── */
async function testPrev(){
  console.log('\n[PREV] v1.5 재실행: debug 플래그 sessionStorage 유지');
  const domA=new JSDOM(
    '<!DOCTYPE html><html data-theme="dark"><head></head><body><div id="app" style="display:none"></div></body></html>',
    {url:'https://mhlee205.github.io/TEST/?debug=1',runScripts:'dangerously',pretendToBeVisual:true}
  );
  const wA=domA.window;
  def(wA,'crypto',{subtle:{digest:async()=>new ArrayBuffer(32)},getRandomValues:(a)=>{for(let i=0;i<a.length;i++)a[i]=i%256;return a;}});
  const ssA={_d:{},getItem(k){return this._d[k]||null;},setItem(k,v){this._d[k]=v;},removeItem(k){delete this._d[k];}};
  def(wA,'sessionStorage',ssA);
  def(wA,'localStorage',{_d:{},getItem(k){return this._d[k]||null;},setItem(k,v){this._d[k]=v;},removeItem(k){delete this._d[k];}});
  def(wA,'history',{replaceState:()=>{}});
  def(wA,'requestAnimationFrame',(cb)=>setTimeout(cb,0));
  def(wA,'fetch',async()=>({ok:false,status:500,json:async()=>({error:'x'})}));
  wA.eval(src);
  await wA.LEENAI.init({sys:'Z-05',name:'T',version:'1',date:'2026-10-08',onReady:()=>{}});
  await new Promise(r=>setTimeout(r,30));
  ok('[PREV] debug 플래그 저장 후 유지',ssA._d['leenai_debug_Z-05']==='1');
}

(async()=>{
  try{await test1();}catch(e){console.error('TEST1 예외:',e.message);failed++;}
  try{await test2();}catch(e){console.error('TEST2 예외:',e.message);failed++;}
  try{await test3();}catch(e){console.error('TEST3 예외:',e.message);failed++;}
  try{await test4();}catch(e){console.error('TEST4 예외:',e.message);failed++;}
  try{await test5();}catch(e){console.error('TEST5 예외:',e.message);failed++;}
  try{await testPrev();}catch(e){console.error('PREV 예외:',e.message);failed++;}
  console.log('\n─────────────────────────');
  console.log('결과: PASS '+passed+' / FAIL '+failed);
  process.exit(failed>0?1:0);
})();
