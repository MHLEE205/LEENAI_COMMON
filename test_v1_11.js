/* leenai-common.js v1.11 jsdom 시험 — init 옵션 graphScopes (하위 호환)
   TEST1: 옵션 없음 → authorize·token 요청이 v1.10 과 글자 그대로 같음 (기존 TOOL 회귀 없음)
   TEST2: 옵션 있음 → authorize 에 추가 권한, 첫 Graph 토큰 요청 = Mail.ReadWrite 묶음 + 추가 권한, graphExtra=ok
   TEST3: 옵션 있음 + 추가 권한 거부 → 기존 Mail.ReadWrite 묶음으로 내려가 Graph ok, graphExtra=error:consent_required
   실행: NODE_PATH=<jsdom 설치 폴더>/node_modules node test_v1_11.js */
'use strict';
const {JSDOM}=require('jsdom');
const fs=require('fs');
const path=require('path');
const src=fs.readFileSync(path.join(__dirname,'v1','leenai-common.js'),'utf8');
const srcOld=process.env.OLD_SRC?fs.readFileSync(process.env.OLD_SRC,'utf8'):null;

function fakeIdToken(name,email){
  const b64=Buffer.from(JSON.stringify({name,preferred_username:email,email}),'utf8').toString('base64url');
  return 'hdr.'+b64+'.sig';
}
function def(w,k,v){Object.defineProperty(w,k,{value:v,writable:true,configurable:true});}

function makeDOM(withCode,calls,rejectExtra){
  const dom=new JSDOM('<!DOCTYPE html><html data-theme="dark"><head></head><body><div id="app" style="display:none"></div></body></html>',
    {url:'https://mhlee205.github.io/CONTRACT_SEND/mail/'+(withCode?'?code=C&state=S':''),runScripts:'dangerously',pretendToBeVisual:true});
  const w=dom.window;
  def(w,'crypto',{subtle:{digest:async()=>new ArrayBuffer(32)},getRandomValues:(a)=>{for(let i=0;i<a.length;i++)a[i]=i%256;return a;}});
  const ss={_d:{'leenai_pkce_MAIL':JSON.stringify({verifier:'v',state:'S'})},getItem(k){return this._d[k]||null;},setItem(k,v){this._d[k]=v;},removeItem(k){delete this._d[k];}};
  def(w,'sessionStorage',ss);
  def(w,'localStorage',{_d:{},getItem(k){return this._d[k]||null;},setItem(k,v){this._d[k]=v;},removeItem(k){delete this._d[k];}});
  def(w,'history',{replaceState:()=>{}});
  def(w,'requestAnimationFrame',(cb)=>setTimeout(cb,0));
  def(w,'fetch',async function(url,opts){
    const body=(opts&&opts.body)||'';
    calls.push(decodeURIComponent(body));
    const json=(o,ok=true)=>({ok,status:ok?200:400,json:async()=>o});
    if(body.includes('authorization_code')) return json({access_token:'DV0',expires_in:3600,refresh_token:'RT0',id_token:fakeIdToken('T','t@t.com')});
    if(body.includes('Sites.ReadWrite.All')&&rejectExtra) return json({error:'consent_required'},false);
    if(body.includes('refresh_token')) return json({access_token:'GW0',expires_in:3600,refresh_token:'RT1'});
    return json({},false);
  });
  return dom;
}

let passed=0,failed=0;
function ok(label,cond){if(cond){console.log('  ✔ '+label);passed++;}else{console.error('  ✖ FAIL: '+label);failed++;}}
const wait=(ms)=>new Promise(r=>setTimeout(r,ms));

async function loginCalls(code,opts,reject,source){
  const calls=[];const dom=makeDOM(true,calls,reject);
  dom.window.eval(code);
  const L=dom.window.LEENAI;
  await L.init(Object.assign({sys:'MAIL',name:'T',version:'1',date:'2026-10-20',onReady:()=>{}},opts));
  await wait(80);
  return {calls,L,dom};
}
async function authorizeUrl(code,opts){
  const dom=makeDOM(false,[],false);
  let href='';
  dom.window.eval(code);
  const L=dom.window.LEENAI;
  await L.init(Object.assign({sys:'MAIL',name:'T',version:'1',date:'2026-10-20',onReady:()=>{}},opts));
  // 로그인 버튼 = splash 의 버튼
  const btn=[...dom.window.document.querySelectorAll('button')].find(b=>/ログイン|Microsoft/.test(b.textContent));
  Object.defineProperty(dom.window,'location',{value:{get href(){return href;},set href(v){href=v;},origin:'https://mhlee205.github.io',pathname:'/CONTRACT_SEND/mail/'},configurable:true});
  if(btn) btn.click();
  await wait(50);
  return decodeURIComponent(href);
}

(async()=>{
  console.log('\n[TEST 1] 옵션 없음 → v1.10 과 같은 요청');
  const a=await loginCalls(src,{},false);
  ok('VERSION v1.11',a.L.VERSION==='v1.11');
  ok('Graph ok',a.L.auth.status.graph==='ok');
  ok('Sites 권한 요청 없음',!a.calls.some(c=>c.includes('Sites.ReadWrite.All')));
  ok('graphExtra 없음',a.L.auth.status.graphExtra===undefined);
  if(srcOld){
    const b=await loginCalls(srcOld,{},false);
    ok('token 요청 순서·본문이 v1.10 과 완전히 같음',JSON.stringify(a.calls)===JSON.stringify(b.calls));
  }
  console.log('\n[TEST 2] graphScopes 지정');
  const c=await loginCalls(src,{graphScopes:['Mail.ReadWrite','Sites.ReadWrite.All']},false);
  const g=c.calls.filter(x=>x.includes('grant_type=refresh_token'));
  ok('첫 Graph 요청 = 묶음 + Sites.ReadWrite.All',g[0].includes('https://graph.microsoft.com/Sites.ReadWrite.All')&&g[0].includes('Mail.Send'));
  ok('Mail.ReadWrite 중복 없음',(g[0].match(/Mail\.ReadWrite/g)||[]).length===1);
  ok('graphExtra=ok',c.L.auth.status.graphExtra==='ok');
  ok('Graph token',await c.L.auth.token('graph')==='GW0');
  console.log('\n[TEST 3] 추가 권한 거부 → 기존 흐름');
  const d=await loginCalls(src,{graphScopes:['Sites.ReadWrite.All']},true);
  ok('Graph ok (기존 묶음으로)',d.L.auth.status.graph==='ok');
  ok('graphExtra=error:consent_required',d.L.auth.status.graphExtra==='error:consent_required');
  const dg=d.calls.filter(x=>x.includes('grant_type=refresh_token'));
  ok('두 번째 요청은 v1.10 의 첫 요청과 같음 (Sites 없음)',!dg[1].includes('Sites.ReadWrite.All')&&dg[1].includes('Mail.ReadWrite'));
  console.log('\n결과: '+passed+' passed / '+failed+' failed');
  process.exit(failed?1:0);
})().catch(e=>{console.error(e);process.exit(1);});
