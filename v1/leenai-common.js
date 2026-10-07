/* LEENAI_COMMON v1.1 (2026-10-07) — DOC-Z-05 v1.0
   window.LEENAI を 1 つだけ公開。他のグローバル変数は作らない。
   ChangeLog: v1.0 — 新規
              v1.1 — id_token UTF-8デコード修正(TextDecoder), staffCode をメール prefix 基準に変更, 未登録者は空文字 */
(function(){
'use strict';

/* ── 定数 ── */
const CLIENT_ID = 'f071a165-5e9b-44bf-b6a1-baba654524db';
const TENANT_ID = 'a56988ff-c2b8-4df8-9727-889ad4205198';
const DV_URL    = 'https://orgde512c6f.crm7.dynamics.com';
const AUTH_URL  = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0`;

const THEMES = ['dark','light','leenai','focus'];
const THEME_LABEL = {dark:'🌙 ダーク',light:'☀️ ライト',leenai:'🌿 LEENAI',focus:'🎯 FOCUS'};

const DEFAULT_STAFF = [
  {name:'LEE MYEONGHOE', email:'mh_lee@leenearcorp.com',      staffCode:'MH_LEE'},
  {name:'KAIXIN ZHAN',   email:'kx_zhan@leenearcorp.com',     staffCode:'KX_ZHAN'},
  {name:'HINA SUZUKI',   email:'h_suzuki@leenearcorp.com',    staffCode:'H_SUZUKI'},
  {name:'YUKI TANAKA',   email:'y_tanaka@leenearcorp.com',    staffCode:'Y_TANAKA'},
  {name:'TAKESHI MORI',  email:'t_mori@leenearcorp.com',      staffCode:'T_MORI'},
  {name:'SAKURA ITO',    email:'s_ito@leenearcorp.com',       staffCode:'S_ITO'},
  {name:'KENJI NAKAMURA',email:'k_nakamura@leenearcorp.com',  staffCode:'K_NAKAMURA'},
];

/* ── 内部ステート ── */
let _cfg = {};
let _dvToken = null, _dvExpiry = 0;
let _grToken = null, _grExpiry = 0;
let _user = null;
let _authStatus = {dataverse:'error', graph:'skip', graphError:''};
let _headerBtns = [];

/* ── ユーティリティ ── */
function b64url(buf){return btoa(String.fromCharCode(...new Uint8Array(buf))).replace(/\+/g,'-').replace(/\//g,'_').replace(/=/g,'');}
async function sha256(s){return b64url(await crypto.subtle.digest('SHA-256',new TextEncoder().encode(s)));}
function randomStr(n){const a='ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-._~';return Array.from(crypto.getRandomValues(new Uint8Array(n)),v=>a[v%a.length]).join('');}

/* ── テーマ ── */
function _applyTheme(t){
  document.documentElement.setAttribute('data-theme',t);
  document.querySelectorAll('[data-ln="themeBtn"]').forEach(b=>b.textContent=THEME_LABEL[t]);
  _setFavicon();
}
function _toggleTheme(){
  const cur=document.documentElement.getAttribute('data-theme')||'dark';
  const next=THEMES[(THEMES.indexOf(cur)+1)%THEMES.length];
  try{localStorage.setItem('leenai_theme',next);}catch(e){}
  _applyTheme(next);
}
function _setFavicon(){
  const acc=(getComputedStyle(document.documentElement).getPropertyValue('--accent')||'#4f8ef7').trim();
  const svg=`<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 32 32"><rect width="32" height="32" rx="7" fill="#0f1117"/><text x="16" y="22" font-family="Segoe UI,Arial" font-size="15" font-weight="800" text-anchor="middle" fill="#e2e8f0">L<tspan fill="${acc}">A</tspan></text></svg>`;
  let el=document.getElementById('leenaiFavicon');
  if(!el){el=document.createElement('link');el.rel='icon';el.id='leenaiFavicon';document.head.appendChild(el);}
  el.href='data:image/svg+xml,'+encodeURIComponent(svg);
}

/* ── ローディング ── */
let _loadEl=null, _loadMsgEl=null;
function _ensureLoading(){
  if(_loadEl) return;
  _loadEl=document.createElement('div'); _loadEl.className='ln-loading';
  _loadMsgEl=document.createElement('div'); _loadMsgEl.className='ln-loading-msg';
  const sp=document.createElement('div'); sp.className='ln-spinner';
  _loadEl.appendChild(sp); _loadEl.appendChild(_loadMsgEl);
  document.body.appendChild(_loadEl);
}
function _showLoading(msg){_ensureLoading();_loadMsgEl.textContent=msg||'';_loadEl.classList.add('ln-active');}
function _updateLoading(msg){if(_loadMsgEl)_loadMsgEl.textContent=msg;}
function _hideLoading(){if(_loadEl)_loadEl.classList.remove('ln-active');}

/* ── トースト ── */
let _toastWrap=null;
function _ensureToast(){
  if(_toastWrap) return;
  _toastWrap=document.createElement('div'); _toastWrap.className='ln-toast-wrap';
  document.body.appendChild(_toastWrap);
}
function _toast(msg, type){
  type=type||'info';
  _ensureToast();
  const el=document.createElement('div'); el.className=`ln-toast ln-${type}`; el.textContent=msg;
  _toastWrap.appendChild(el);
  requestAnimationFrame(function(){requestAnimationFrame(function(){el.classList.add('ln-in');});});
  setTimeout(function(){el.classList.remove('ln-in');setTimeout(function(){el.remove();},300);},3500);
}

/* ── DOM 構築 ── */
function _buildSplash(){
  const d=document.createElement('div'); d.id='ln-splash'; d.className='ln-splash';
  d.innerHTML='<div class="ln-splash-card">'
    +'<div class="ln-logo">LEEN<span>AI</span></div>'
    +'<div class="ln-splash-meta">'
    +'<span class="ln-sys" data-ln="sys">'+_cfg.sys+'</span>'
    +'<span class="ln-splash-ver" data-ln="ver">v'+_cfg.version+' \xb7 '+_cfg.date+'</span>'
    +'</div>'
    +'<div class="ln-splash-name" data-ln="name">'+_cfg.name+'</div>'
    +'<div class="ln-splash-desc" data-ln="desc">'+(_cfg.desc||'')+'</div>'
    +'<hr>'
    +'<button class="ln-login-btn" id="ln-login-btn">'
    +'<span class="ln-ms"><i style="background:#f25022"></i><i style="background:#7fba00"></i><i style="background:#00a4ef"></i><i style="background:#ffb900"></i></span>'
    +' Microsoft 365 でログイン'
    +'</button>'
    +'<div class="ln-login-note">COMPASS と Outlook に同時にサインインします</div>'
    +'<div class="ln-splash-theme"><button class="ln-btn-ghost" data-ln="themeBtn" onclick="LEENAI._toggleTheme()">🌙 ダーク</button></div>'
    +'</div>'
    +'<div class="ln-footer">❆ Powered by LEENAI Automation System</div>';
  document.body.insertBefore(d, document.body.firstChild);
  document.getElementById('ln-login-btn').addEventListener('click', _startLogin);
}

function _buildHeader(){
  const h=document.createElement('header'); h.className='ln-header'; h.id='ln-header';
  h.innerHTML='<div class="ln-logo">LEEN<span>AI</span></div>'
    +'<span class="ln-sys">'+_cfg.sys+'</span>'
    +'<div class="ln-hdr-sep"></div>'
    +'<span class="ln-hdr-name">'+_cfg.name+'</span>'
    +'<span class="ln-hdr-ver">v'+_cfg.version+' \xb7 '+_cfg.date+'</span>'
    +'<div class="ln-hdr-right" id="ln-hdr-right">'
    +'<span class="ln-user" id="ln-user-chip">'
    +'<span class="ln-avatar" id="ln-avatar"></span>'
    +'<span id="ln-username"></span>'
    +'<span class="ln-dot" id="ln-dot" title="接続中"></span>'
    +'</span>'
    +'<span id="ln-hdr-extra-btns"></span>'
    +'<button class="ln-btn-ghost" data-ln="themeBtn" onclick="LEENAI._toggleTheme()">🌙 ダーク</button>'
    +'<button class="ln-btn-ghost" onclick="LEENAI.logout()">ログアウト</button>'
    +'</div>';
  return h;
}

function _buildFooter(){
  const f=document.createElement('div'); f.className='ln-footer';
  f.textContent='❆ Powered by LEENAI Automation System';
  return f;
}

/* ── ヘッダードット更新 ── */
function _updateDot(){
  const dot=document.getElementById('ln-dot'); if(!dot) return;
  const s=_authStatus;
  if(s.dataverse==='ok'&&(s.graph==='ok'||s.graph==='skip')) dot.style.background='var(--ok)';
  else if(s.dataverse==='ok'&&s.graph==='error') dot.style.background='var(--warn)';
  else dot.style.background='var(--danger)';
}

/* ── ログイン (PKCE) ── */
async function _startLogin(){
  const verifier=randomStr(64);
  const challenge=await sha256(verifier);
  const state=randomStr(32);
  const pkceKey='leenai_pkce_'+_cfg.sys;
  try{sessionStorage.setItem(pkceKey,JSON.stringify({verifier,state}));}catch(e){}
  const rdUri=location.origin+location.pathname.replace(/index\.html$/,'');
  const scope=encodeURIComponent(DV_URL+'/user_impersonation openid profile email offline_access');
  const params='client_id='+CLIENT_ID+'&response_type=code&redirect_uri='+encodeURIComponent(rdUri)+'&scope='+scope+'&code_challenge='+challenge+'&code_challenge_method=S256&state='+state+'&prompt=select_account';
  location.href=AUTH_URL+'/authorize?'+params;
}

async function _handleCallback(){
  const url=new URL(location.href);
  const code=url.searchParams.get('code');
  const retState=url.searchParams.get('state');
  if(!code) return false;
  const pkceKey='leenai_pkce_'+_cfg.sys;
  let pkce;
  try{pkce=JSON.parse(sessionStorage.getItem(pkceKey)||'null');}catch(e){}
  sessionStorage.removeItem(pkceKey);
  if(!pkce||pkce.state!==retState){
    _toast('認証エラー: state 不一致。再ログインしてください。','danger');
    history.replaceState({},'',location.pathname);
    return false;
  }
  history.replaceState({},'',location.pathname);
  const rdUri=location.origin+location.pathname.replace(/index\.html$/,'');
  _showLoading('認証中…');
  try{
    const r1=await fetch(AUTH_URL+'/token',{method:'POST',headers:{'Content-Type':'application/x-www-form-urlencoded'},
      body:'client_id='+CLIENT_ID+'&grant_type=authorization_code&code='+encodeURIComponent(code)+'&redirect_uri='+encodeURIComponent(rdUri)+'&code_verifier='+pkce.verifier+'&scope='+encodeURIComponent(DV_URL+'/user_impersonation openid profile email offline_access')});
    if(!r1.ok) throw new Error('DV token '+r1.status);
    const t1=await r1.json();
    _dvToken=t1.access_token; _dvExpiry=Date.now()+(t1.expires_in-60)*1000;
    _authStatus.dataverse='ok';
    try{
      const _raw=t1.id_token.split('.')[1].replace(/-/g,'+').replace(/_/g,'/');
      const _bytes=Uint8Array.from(atob(_raw),function(c){return c.charCodeAt(0);});
      const payload=JSON.parse(new TextDecoder().decode(_bytes));
      const email=payload.preferred_username||payload.email||'';
      const staff=DEFAULT_STAFF.find(function(s){return s.email.toLowerCase()===email.toLowerCase();});
      const initials=(payload.name||'').split(' ').map(function(w){return w[0];}).join('').slice(0,2).toUpperCase()||'??';
      _user={name:payload.name||email,email:email,staffCode:staff?staff.staffCode:'',initials:initials};
    }catch(e){_user={name:'(不明)',email:'',staffCode:'??',initials:'??'};}
    if(_cfg.needGraph!==false){
      try{
        _updateLoading('Outlook 接続中…');
        const r2=await fetch(AUTH_URL+'/token',{method:'POST',headers:{'Content-Type':'application/x-www-form-urlencoded'},
          body:'client_id='+CLIENT_ID+'&grant_type=refresh_token&refresh_token='+encodeURIComponent(t1.refresh_token)+'&scope='+encodeURIComponent('https://graph.microsoft.com/Mail.Send https://graph.microsoft.com/User.Read offline_access')});
        if(!r2.ok){const e=await r2.json();throw Object.assign(new Error(r2.status),{code:e.error});}
        const t2=await r2.json();
        _grToken=t2.access_token; _grExpiry=Date.now()+(t2.expires_in-60)*1000;
        _authStatus.graph='ok'; _authStatus.graphError='';
      }catch(e){
        _authStatus.graph='error'; _authStatus.graphError=e.code||String(e);
      }
    } else {_authStatus.graph='skip';}
    return true;
  }catch(e){
    _toast('ログイン失敗: '+e.message,'danger');
    _hideLoading();
    return false;
  }
}

function _showSplash(){
  const splash=document.getElementById('ln-splash');
  const app=document.getElementById('app');
  if(splash) splash.style.display='flex';
  if(app) app.style.display='none';
  _dvToken=null;_dvExpiry=0;_grToken=null;_grExpiry=0;_user=null;
  _authStatus={dataverse:'error',graph:_cfg.needGraph===false?'skip':'error',graphError:''};
}

function _showApp(){
  const splash=document.getElementById('ln-splash');
  const app=document.getElementById('app');
  if(splash) splash.style.display='none';
  if(app) app.style.display='';
  if(!document.getElementById('ln-header')){
    const h=_buildHeader();
    if(app) app.insertBefore(h, app.firstChild);
  }
  if(app&&!app.querySelector('.ln-footer')) app.appendChild(_buildFooter());
  const av=document.getElementById('ln-avatar');
  const un=document.getElementById('ln-username');
  if(av&&_user) av.textContent=_user.initials||_user.staffCode;
  if(un&&_user) un.textContent=_user.name;
  const extra=document.getElementById('ln-hdr-extra-btns');
  if(extra){
    extra.innerHTML='';
    _headerBtns.forEach(function(b){
      const el=document.createElement('button'); el.className='ln-btn-ghost'; el.textContent=b.label;
      if(b.title) el.title=b.title;
      el.addEventListener('click',b.onClick); extra.appendChild(el);
    });
  }
  _updateDot();
  let t='dark'; try{t=localStorage.getItem('leenai_theme')||'dark';}catch(e){}
  _applyTheme(t);
}

/* ── init ── */
async function _init(cfg){
  _cfg=Object.assign({needGraph:true},cfg);
  document.title='LEENAI '+_cfg.sys+' '+_cfg.name+' v'+_cfg.version;
  Array.from(document.body.children).forEach(function(el){if(!el.id||el.id!=='ln-splash') el.style.display='none';});
  let t='dark'; try{t=localStorage.getItem('leenai_theme')||'dark';}catch(e){}
  _applyTheme(t);
  _buildSplash();
  const url=new URL(location.href);
  if(url.searchParams.has('code')){
    const ok=await _handleCallback();
    if(ok){
      _hideLoading();
      _showApp();
      if(typeof _cfg.onReady==='function') _cfg.onReady();
    } else {_showSplash();}
  } else {_showSplash();}
}

/* ── logout ── */
function _logout(){
  _dvToken=null;_dvExpiry=0;_grToken=null;_grExpiry=0;_user=null;
  _authStatus={dataverse:'error',graph:_cfg.needGraph===false?'skip':'error',graphError:''};
  _showSplash();
}

/* ── addHeaderButton ── */
function _addHeaderButton(opts){
  _headerBtns.push(opts);
  const extra=document.getElementById('ln-hdr-extra-btns');
  if(extra){
    const el=document.createElement('button'); el.className='ln-btn-ghost'; el.textContent=opts.label;
    if(opts.title) el.title=opts.title;
    el.addEventListener('click',opts.onClick); extra.insertBefore(el,extra.firstChild);
  }
}

/* ── トークン取得 ── */
async function _getToken(type){
  if(type==='dataverse'){
    if(_dvToken&&Date.now()<_dvExpiry) return _dvToken;
    _toast('セッションが切れました。再ログインしてください。','warn');
    setTimeout(function(){_showSplash();},2000);
    throw new Error('session expired');
  }
  if(type==='graph'){
    if(_authStatus.graph==='skip') return null;
    if(_grToken&&Date.now()<_grExpiry) return _grToken;
    _authStatus.graph='error'; _updateDot(); return null;
  }
  throw new Error('unknown token type: '+type);
}

/* ── 公開 API ── */
window.LEENAI={
  init:_init,
  auth:{
    get status(){return Object.assign({},_authStatus);},
    token:_getToken,
  },
  get user(){return _user?Object.assign({},_user):null;},
  logout:_logout,
  loading:{show:_showLoading,update:_updateLoading,hide:_hideLoading},
  toast:_toast,
  addHeaderButton:_addHeaderButton,
  STAFF:DEFAULT_STAFF,
  COMPANY:{tel:'+81-3-3528-9850',fax:'+81-3-3528-9851'},
  VERSION:'v1.1',
  _toggleTheme:_toggleTheme,
};
})();
