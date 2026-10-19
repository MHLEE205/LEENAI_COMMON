/* leenai-common.js v1.10 jsdom 시험 — 마스터 ref 칸 (드롭다운)
   TEST1: ref 칸이 <select> 로 그려지고, 선택지는 직원 마스터의 이름 (在籍 해제 제외)
   TEST2: 화면 값 변경 → 저장 시 코드로 저장, 마스터에 없는 값은 저장 거부
   TEST3: LEENAI.master.label() 코드 → 이름
   TEST4: 기존 text·bool 칸은 그대로 <input>
   실행: NODE_PATH=<jsdom 설치 폴더>/node_modules node test_v1_10.js */
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

const STAFF={master:'staff',fields:[{key:'code',label:'STAFF_CODE',type:'text',required:true},{key:'name_ja',label:'名前',type:'text'},{key:'active',label:'在籍',type:'bool'}],
  unique:['code'],items:[{code:'M_TAKEUCHI',name_ja:'竹内 雅将',active:true},{code:'SM_PARK',name_ja:'朴 秀旻',active:true},{code:'OLD_MAN',name_ja:'退職 太郎',active:false}]};
const POL={master:'pol_staff',title:'POL 担当',fields:[{key:'pol',label:'POL',type:'text',required:true},
  {key:'staff',label:'担当',type:'ref',ref:'staff',ref_key:'code',ref_label:'name_ja',required:true},
  {key:'assist',label:'補助',type:'ref',ref:'staff',ref_key:'code',ref_label:'name_ja'}],
  unique:['pol'],history:[],items:[{pol:'TOKYO',staff:'M_TAKEUCHI',assist:'SM_PARK'}]};

let puts=[];
function makeDOM(){
  const dom=new JSDOM('<!DOCTYPE html><html data-theme="dark"><head></head><body><div id="app" style="display:none"></div></body></html>',
    {url:'https://mhlee205.github.io/TEST/?debug=1&code=C&state=S',runScripts:'dangerously',pretendToBeVisual:true});
  const w=dom.window;
  def(w,'crypto',{subtle:{digest:async()=>new ArrayBuffer(32)},getRandomValues:(a)=>{for(let i=0;i<a.length;i++)a[i]=i%256;return a;}});
  const ss={_d:{'leenai_pkce_Z-05':JSON.stringify({verifier:'v',state:'S'})},getItem(k){return this._d[k]||null;},setItem(k,v){this._d[k]=v;},removeItem(k){delete this._d[k];}};
  def(w,'sessionStorage',ss);
  def(w,'localStorage',{_d:{},getItem(k){return this._d[k]||null;},setItem(k,v){this._d[k]=v;},removeItem(k){delete this._d[k];}});
  def(w,'history',{replaceState:()=>{}});
  def(w,'requestAnimationFrame',(cb)=>setTimeout(cb,0));
  def(w,'confirm',()=>true);
  def(w,'fetch',async function(url,opts){
    const body=opts&&opts.body||'';
    const json=(o)=>({ok:true,status:200,headers:{get:()=>'"etag1"'},json:async()=>o});
    if(typeof body==='string'&&body.includes('authorization_code'))
      return json({access_token:'DV0',expires_in:3600,refresh_token:'RT0',id_token:fakeIdToken('TEST','t@t.com')});
    if(typeof body==='string'&&body.includes('refresh_token'))
      return json({access_token:'GW0',expires_in:3600,refresh_token:'RT1'});
    const u=decodeURIComponent(url);
    if(opts&&opts.method==='PUT'){puts.push(JSON.parse(body));return json({'@odata.etag':'"etag2"'});}
    if(u.includes(':/children')) return json({value:[{name:'master_pol_staff.json'},{name:'master_staff.json'}]});
    if(u.includes('master_staff.json')) return json(STAFF);
    if(u.includes('master_pol_staff.json')) return json(POL);
    if(u.endsWith('MASTER')) return json({webUrl:'https://x'});
    return {ok:false,status:404,headers:{get:()=>''},json:async()=>({})};
  });
  return dom;
}

let passed=0,failed=0;
function ok(label,cond){if(cond){console.log('  ✔ '+label);passed++;}else{console.error('  ✖ FAIL: '+label);failed++;}}
const wait=(ms)=>new Promise(r=>setTimeout(r,ms));

async function main(){
  const dom=makeDOM(); const w=dom.window; const d=w.document;
  w.eval(src);
  const L=w.LEENAI;
  await L.init({sys:'Z-05',name:'T',version:'1',date:'2026-10-19',onReady:()=>{}});
  await wait(80);
  ok('VERSION v1.10',L.VERSION==='v1.10');

  console.log('\n[TEST 1] ref 칸 → <select>');
  L._openMasterPanel(); await wait(80);
  const item=[...d.querySelectorAll('.ln-master-sitem')].find(e=>e.getAttribute('data-mname')==='pol_staff');
  item.click(); await wait(80);
  const sels=d.querySelectorAll('#ln-master-tbody select');
  ok('select 2개 (担当·補助)',sels.length===2);
  const opts=[...sels[0].options].map(o=>o.textContent);
  ok('이름으로 표시',opts.includes('竹内 雅将')&&opts.includes('朴 秀旻'));
  ok('在籍 해제 제외',!opts.some(t=>t.includes('退職')));
  ok('현재 값 선택됨',sels[0].value==='M_TAKEUCHI');
  ok('POL 칸은 input',d.querySelector('#ln-master-tbody input[data-key="pol"]')!==null);

  console.log('\n[TEST 2] 저장');
  sels[1].value='M_TAKEUCHI'; sels[1].dispatchEvent(new w.Event('change'));
  await L._masterSave(); await wait(50);
  ok('PUT 1회',puts.length===1);
  ok('코드로 저장',puts[0].items[0].assist==='M_TAKEUCHI');
  // 없는 값
  await wait(50);
  const s2=d.querySelectorAll('#ln-master-tbody select');
  const o=d.createElement('option'); o.value='NOBODY'; o.textContent='x'; s2[0].appendChild(o);
  s2[0].value='NOBODY'; s2[0].dispatchEvent(new w.Event('change'));
  await L._masterSave(); await wait(50);
  ok('마스터에 없는 값 → 저장 거부',puts.length===1&&d.getElementById('ln-master-err').textContent.includes('マスターにない値'));

  console.log('\n[TEST 3] master.label');
  ok('label → 이름',L.master.label('pol_staff','staff','SM_PARK')==='朴 秀旻');
  ok('ref 아닌 칸은 그대로',L.master.label('pol_staff','pol','TOKYO')==='TOKYO');

  console.log('\n[TEST 4] 기존 칸 종류 그대로');
  const st=[...d.querySelectorAll('.ln-master-sitem')].find(e=>e.getAttribute('data-mname')==='staff');
  st.click(); await wait(80);
  ok('staff 마스터에는 select 없음',d.querySelectorAll('#ln-master-tbody select').length===0);
  ok('bool 은 checkbox',d.querySelectorAll('#ln-master-tbody input[type=checkbox]').length===3);

  console.log('\n결과: '+passed+' passed / '+failed+' failed');
  process.exit(failed?1:0);
}
main().catch(e=>{console.error(e);process.exit(1);});
