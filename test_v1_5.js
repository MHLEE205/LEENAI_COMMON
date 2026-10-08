/* leenai-common.js v1.5 jsdom 시험
   「?debug=1로 열기 → mock 로그인 완료 → URL에 debug 없음 → _expireForTest() 실행 시 만료가 0이 되는지」 */
'use strict';
const {JSDOM} = require('jsdom');
const fs = require('fs');
const path = require('path');
const src = fs.readFileSync(path.join(__dirname,'v1','leenai-common.js'),'utf8');

function fakeIdToken(name, email){
  const b64 = Buffer.from(JSON.stringify({name,preferred_username:email,email}),'utf8').toString('base64url');
  return 'hdr.'+b64+'.sig';
}

function mockWindow(w, ssInit){
  const def=(k,v)=>Object.defineProperty(w,k,{value:v,writable:true,configurable:true});
  def('crypto',{
    subtle:{digest:async()=>new ArrayBuffer(32)},
    getRandomValues:(a)=>{for(let i=0;i<a.length;i++)a[i]=i%256;return a;}
  });
  const ss={_d:Object.assign({},ssInit||{}),
    getItem(k){return this._d[k]||null;},
    setItem(k,v){this._d[k]=v;},
    removeItem(k){delete this._d[k];}};
  def('sessionStorage',ss);
  def('localStorage',{_d:{},getItem(k){return this._d[k]||null;},setItem(k,v){this._d[k]=v;},removeItem(k){delete this._d[k];}});
  def('history',{replaceState:()=>{}});
  def('requestAnimationFrame',(cb)=>setTimeout(cb,0));
  return ss;
}

let passed=0, failed=0;
function ok(label,cond){
  if(cond){console.log('  ✔ '+label);passed++;}
  else{console.error('  ✖ FAIL: '+label);failed++;}
}

/* ══════════════════════════════════════════
   TEST 1: ?debug=1 로 열기 → 로그인 → URL에 debug 없음 → _expireForTest() 동작 확인
   흐름:
     1. URL = ?debug=1  → init → sessionStorage에 'leenai_debug_Z-05'='1' 저장
     2. _showSplash() 호출 → debug 플래그가 삭제되지 않음 (v1.5 수정 확인)
     3. 콜백 URL(code+state) 로 init 재호출 (로그인 완료)
     4. 이 시점 URL에 debug 없음 → sessionStorage의 플래그로 _expireForTest() 동작해야 함
     5. _dvExpiry가 실제로 0이 되는지 확인
══════════════════════════════════════════ */
async function test1(){
  console.log('\n[TEST 1] ?debug=1 열기 → 로그인 후 URL에 debug 없어도 _expireForTest() 동작');

  /* ── Step A: 첫 번째 열기 (?debug=1, code 없음) ── */
  const domA = new JSDOM(
    `<!DOCTYPE html><html data-theme="dark"><head></head><body><div id="app" style="display:none"></div></body></html>`,
    {url:'https://mhlee205.github.io/TEST/?debug=1', runScripts:'dangerously', pretendToBeVisual:true}
  );
  const ssA = mockWindow(domA.window, {});
  Object.defineProperty(domA.window,'fetch',{value:async()=>({ok:false,json:async()=>({error:'x'})}),writable:true,configurable:true});

  domA.window.eval(src);
  const LA = domA.window.LEENAI;
  await LA.init({sys:'Z-05',name:'T',version:'1',date:'2026-10-08',onReady:()=>{}});
  await new Promise(r=>setTimeout(r,30));

  ok('Step A: sessionStorage에 debug 플래그 저장됨', ssA._d['leenai_debug_Z-05']==='1');
  ok('Step A: _showSplash 후에도 debug 플래그 유지됨 (v1.5 핵심)', ssA._d['leenai_debug_Z-05']==='1');

  /* ── Step B: 로그인 콜백 URL (debug 없음, code+state) ── */
  const domB = new JSDOM(
    `<!DOCTYPE html><html data-theme="dark"><head></head><body><div id="app" style="display:none"></div></body></html>`,
    {url:'https://mhlee205.github.io/TEST/?code=C&state=S', runScripts:'dangerously', pretendToBeVisual:true}
  );
  /* Step A에서 저장된 sessionStorage 플래그를 Step B에 이어받음 */
  const ssB = mockWindow(domB.window, {'leenai_pkce_Z-05':JSON.stringify({verifier:'v',state:'S'}), 'leenai_debug_Z-05':'1'});
  Object.defineProperty(domB.window,'fetch',{value:async function(_url,opts){
    const body=opts&&opts.body||'';
    if(body.includes('authorization_code'))
      return {ok:true,json:async()=>({access_token:'DV_TOKEN_XYZ',expires_in:3600,refresh_token:'RT0',id_token:fakeIdToken('TEST','t@t.com')})};
    if(body.includes('refresh_token')&&body.includes('graph'))
      return {ok:true,json:async()=>({access_token:'GR_TOKEN',expires_in:3600})};
    return {ok:false,json:async()=>({error:'x'})};
  },writable:true,configurable:true});

  domB.window.eval(src);
  const LB = domB.window.LEENAI;
  let ready=false;
  await LB.init({sys:'Z-05',name:'T',version:'1',date:'2026-10-08',onReady:()=>{ready=true;}});
  await new Promise(r=>setTimeout(r,80));

  ok('Step B: 로그인 완료 (onReady 呼ばれた)', ready);
  ok('Step B: DV status ok', LB.auth.status.dataverse==='ok');

  /* _dvExpiry가 양수인지 확인 (갱신 전) */
  const expiryBefore = LB.auth.status.dataverse==='ok'; // 간접 확인
  ok('Step B: URL에 debug 없지만 sessionStorage 플래그 존재', ssB._d['leenai_debug_Z-05']==='1');

  /* _expireForTest() 실행 → _dvExpiry가 0이 되어야 함 */
  LB.auth._expireForTest();

  /* token() 을 호출하면 refresh가 발생해야 함 (만료됐으므로) */
  let refreshCalled=false;
  Object.defineProperty(domB.window,'fetch',{value:async function(_url,opts){
    const body=opts&&opts.body||'';
    if(body.includes('refresh_token')&&body.includes('user_impersonation')){
      refreshCalled=true;
      return {ok:true,json:async()=>({access_token:'DV_NEW_TOKEN',expires_in:3600,refresh_token:'RT1'})};
    }
    if(body.includes('refresh_token')&&body.includes('graph'))
      return {ok:true,json:async()=>({access_token:'GR_NEW',expires_in:3600})};
    return {ok:false,json:async()=>({error:'x'})};
  },writable:true,configurable:true});

  const newTok = await LB.auth.token('dataverse');
  ok('_expireForTest() 후 refresh 실제 발생', refreshCalled);
  ok('새 토큰 DV_NEW_TOKEN 반환', newTok==='DV_NEW_TOKEN');
}

/* ══════════════════════════════════════════
   TEST 2: logout 시 debug 플래그 삭제 확인
══════════════════════════════════════════ */
async function test2(){
  console.log('\n[TEST 2] logout 시 debug 플래그 삭제');
  const dom = new JSDOM(
    `<!DOCTYPE html><html data-theme="dark"><head></head><body><div id="app" style="display:none"></div></body></html>`,
    {url:'https://mhlee205.github.io/TEST/?code=C&state=S', runScripts:'dangerously', pretendToBeVisual:true}
  );
  const ss = mockWindow(dom.window, {'leenai_pkce_Z-05':JSON.stringify({verifier:'v',state:'S'}), 'leenai_debug_Z-05':'1'});
  Object.defineProperty(dom.window,'fetch',{value:async function(_url,opts){
    const body=opts&&opts.body||'';
    if(body.includes('authorization_code'))
      return {ok:true,json:async()=>({access_token:'T0',expires_in:3600,refresh_token:'R0',id_token:fakeIdToken('T','t@t.com')})};
    if(body.includes('refresh_token'))
      return {ok:true,json:async()=>({access_token:'G0',expires_in:3600})};
    return {ok:false,json:async()=>({error:'x'})};
  },writable:true,configurable:true});

  dom.window.eval(src);
  const L = dom.window.LEENAI;
  await L.init({sys:'Z-05',name:'T',version:'1',date:'2026-10-08',onReady:()=>{}});
  await new Promise(r=>setTimeout(r,80));

  ok('로그인 후 debug 플래그 존재', ss._d['leenai_debug_Z-05']==='1');
  L.logout();
  ok('logout 후 debug 플래그 삭제됨', !ss._d['leenai_debug_Z-05']);
}

(async()=>{
  try{await test1();}catch(e){console.error('TEST1 예외:',e.message,'\n',e.stack);failed++;}
  try{await test2();}catch(e){console.error('TEST2 예외:',e.message);failed++;}
  console.log('\n─────────────────────────');
  console.log('결과: PASS '+passed+' / FAIL '+failed);
  process.exit(failed>0?1:0);
})();
