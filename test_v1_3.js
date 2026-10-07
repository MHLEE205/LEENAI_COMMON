/* leenai-common.js v1.3 jsdom 자동시험 */
'use strict';
const {JSDOM} = require('jsdom');
const fs = require('fs');
const path = require('path');
const src = fs.readFileSync(path.join(__dirname,'v1','leenai-common.js'),'utf8');

function fakeIdToken(name, email){
  const b64 = Buffer.from(JSON.stringify({name,preferred_username:email,email}),'utf8').toString('base64url');
  return 'hdr.'+b64+'.sig';
}

/* jsdom window に mock を安全に設定 */
function mockWindow(w, pkceState){
  const def = (k,v)=>Object.defineProperty(w,k,{value:v,writable:true,configurable:true});

  def('crypto',{
    subtle:{digest:async()=>new ArrayBuffer(32)},
    getRandomValues:(a)=>{for(let i=0;i<a.length;i++)a[i]=i%256;return a;}
  });

  const ss={_d:{},getItem(k){return this._d[k]||null;},setItem(k,v){this._d[k]=v;},removeItem(k){delete this._d[k];}};
  if(pkceState) ss._d['leenai_pkce_Z-05']=JSON.stringify({verifier:'v',state:pkceState});
  def('sessionStorage', ss);
  def('localStorage',{_d:{},getItem(k){return this._d[k]||null;},setItem(k,v){this._d[k]=v;},removeItem(k){delete this._d[k];}});
  def('history',{replaceState:()=>{}});
  def('requestAnimationFrame',(cb)=>setTimeout(cb,0));
}

function makeDOM(fetchFn, pkceState='S'){
  const dom = new JSDOM(
    `<!DOCTYPE html><html data-theme="dark"><head></head><body><div id="app" style="display:none"></div></body></html>`,
    {url:'https://mhlee205.github.io/TEST/?debug=1&code=C&state=S', runScripts:'dangerously', pretendToBeVisual:true}
  );
  mockWindow(dom.window, pkceState);
  Object.defineProperty(dom.window,'fetch',{value:fetchFn,writable:true,configurable:true});
  return dom;
}

let passed=0, failed=0;
function ok(label,cond){
  if(cond){console.log('  ✔ '+label);passed++;}
  else{console.error('  ✖ FAIL: '+label);failed++;}
}

/* ── TEST 1: 만료 → DV refresh → 새 토큰 ── */
async function test1(){
  console.log('\n[TEST 1] 만료 → DV refresh 성공 → 새 토큰 반환');
  let dvRefreshCount=0;
  const dom = makeDOM(async function(_url,opts){
    const body=opts&&opts.body||'';
    if(body.includes('authorization_code'))
      return {ok:true,json:async()=>({access_token:'OLD_DV',expires_in:3600,refresh_token:'RT0',id_token:fakeIdToken('李明鎬','mh_lee@leenearcorp.com')})};
    if(body.includes('refresh_token')&&body.includes('user_impersonation')){dvRefreshCount++;return {ok:true,json:async()=>({access_token:'NEW_DV',expires_in:3600,refresh_token:'RT1'})};}
    if(body.includes('refresh_token')&&body.includes('graph'))
      return {ok:true,json:async()=>({access_token:'NEW_GR',expires_in:3600})};
    return {ok:false,json:async()=>({error:'x'})};
  });
  dom.window.eval(src);
  const L=dom.window.LEENAI;
  let ready=false;
  await L.init({sys:'Z-05',name:'T',version:'1',date:'2026-10-07',onReady:()=>{ready=true;}});
  await new Promise(r=>setTimeout(r,80));

  ok('onReady 呼ばれた', ready);
  ok('DV status ok', L.auth.status.dataverse==='ok');
  ok('user.name UTF-8正常 (李明鎬)', L.user&&L.user.name==='李明鎬');
  ok('staffCode MH_LEE', L.user&&L.user.staffCode==='MH_LEE');

  L.auth._expireForTest(); dvRefreshCount=0;
  const tok=await L.auth.token('dataverse');
  ok('DV refresh 1回呼ばれた', dvRefreshCount===1);
  ok('新トークン NEW_DV', tok==='NEW_DV');
  ok('DV status still ok', L.auth.status.dataverse==='ok');

  const gr=await L.auth.token('graph');
  ok('Graph token NEW_GR', gr==='NEW_GR');
}

/* ── TEST 2: 병렬 token() → refresh 1회만 ── */
async function test2(){
  console.log('\n[TEST 2] 병렬 token() → DV refresh 1回だけ');
  let refreshCount=0;
  const dom=makeDOM(async function(_url,opts){
    const body=opts&&opts.body||'';
    if(body.includes('authorization_code'))
      return {ok:true,json:async()=>({access_token:'T0',expires_in:3600,refresh_token:'R0',id_token:fakeIdToken('T','t@t.com')})};
    if(body.includes('refresh_token')&&body.includes('user_impersonation')){
      refreshCount++;await new Promise(r=>setTimeout(r,20));
      return {ok:true,json:async()=>({access_token:'T1',expires_in:3600,refresh_token:'R1'})};
    }
    if(body.includes('refresh_token')&&body.includes('graph'))
      return {ok:true,json:async()=>({access_token:'G1',expires_in:3600})};
    return {ok:false,json:async()=>({error:'x'})};
  });
  dom.window.eval(src);
  const L=dom.window.LEENAI;
  await L.init({sys:'Z-05',name:'T',version:'1',date:'2026-10-07',onReady:()=>{}});
  await new Promise(r=>setTimeout(r,80));

  L.auth._expireForTest(); refreshCount=0;
  const [a,b,c]=await Promise.all([L.auth.token('dataverse'),L.auth.token('dataverse'),L.auth.token('dataverse')]);
  ok('refresh 1回だけ', refreshCount===1);
  ok('3つ全部同じ', a===b&&b===c);
  ok('トークン T1', a==='T1');
}

/* ── TEST 3: DV refresh 실패 → throw + status error ── */
async function test3(){
  console.log('\n[TEST 3] DV refresh 실패 → 例外スロー + status.dataverse=error');
  const dom=makeDOM(async function(_url,opts){
    const body=opts&&opts.body||'';
    if(body.includes('authorization_code'))
      return {ok:true,json:async()=>({access_token:'T0',expires_in:3600,refresh_token:'R0',id_token:fakeIdToken('T','t@t.com')})};
    if(body.includes('refresh_token')&&body.includes('user_impersonation'))
      return {ok:false,json:async()=>({error:'invalid_grant'})};
    if(body.includes('refresh_token')&&body.includes('graph'))
      return {ok:true,json:async()=>({access_token:'G0',expires_in:3600})};
    return {ok:false,json:async()=>({error:'x'})};
  });
  Object.defineProperty(dom.window,'setTimeout',{value:(fn,ms)=>setTimeout(fn,Math.min(ms,50)),writable:true,configurable:true});
  dom.window.eval(src);
  const L=dom.window.LEENAI;
  await L.init({sys:'Z-05',name:'T',version:'1',date:'2026-10-07',onReady:()=>{}});
  await new Promise(r=>setTimeout(r,80));

  L.auth._expireForTest();
  let threw=false;
  try{await L.auth.token('dataverse');}catch(e){threw=true;}
  await new Promise(r=>setTimeout(r,100));

  ok('token() が例外スロー', threw);
  ok('status.dataverse=error', L.auth.status.dataverse==='error');
}

/* ── TEST 4: Graph refresh 실패 → DV 유지, graph=error ── */
async function test4(){
  console.log('\n[TEST 4] Graph refresh 실패 → DV 유지, graph=error');
  const dom=makeDOM(async function(_url,opts){
    const body=opts&&opts.body||'';
    if(body.includes('authorization_code'))
      return {ok:true,json:async()=>({access_token:'DV0',expires_in:3600,refresh_token:'R0',id_token:fakeIdToken('T','t@t.com')})};
    if(body.includes('refresh_token')&&body.includes('user_impersonation'))
      return {ok:true,json:async()=>({access_token:'DV1',expires_in:3600,refresh_token:'R1'})};
    if(body.includes('refresh_token')&&body.includes('graph'))
      return {ok:false,json:async()=>({error:'consent_required'})};
    return {ok:false,json:async()=>({error:'x'})};
  });
  dom.window.eval(src);
  const L=dom.window.LEENAI;
  await L.init({sys:'Z-05',name:'T',version:'1',date:'2026-10-07',onReady:()=>{}});
  await new Promise(r=>setTimeout(r,80));

  L.auth._expireForTest();
  const dv=await L.auth.token('dataverse');
  const gr=await L.auth.token('graph');

  ok('DV 갱신 성공 DV1', dv==='DV1');
  ok('Graph null 반환', gr===null);
  ok('DV status ok', L.auth.status.dataverse==='ok');
  ok('Graph status error', L.auth.status.graph==='error');
  ok('graphError consent_required', L.auth.status.graphError==='consent_required');
}

(async()=>{
  try{await test1();}catch(e){console.error('TEST1 예외:',e.message,e.stack);failed++;}
  try{await test2();}catch(e){console.error('TEST2 예외:',e.message);failed++;}
  try{await test3();}catch(e){console.error('TEST3 예외:',e.message);failed++;}
  try{await test4();}catch(e){console.error('TEST4 예외:',e.message);failed++;}
  console.log('\n─────────────────────────');
  console.log('결과: PASS '+passed+' / FAIL '+failed);
  process.exit(failed>0?1:0);
})();
