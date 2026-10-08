/* LEENAI_COMMON v1.9 (2026-10-08) — DOC-Z-05 v1.0
   window.LEENAI を 1 つだけ公開。他のグローバル変数は作らない。
   ChangeLog: v1.0 — 新規
              v1.1 — id_token UTF-8デコード修正(TextDecoder), staffCode をメール prefix 基準に変更, 未登録者は空文字
              v1.2 — バージョン番号修正 (v1.1 → v1.2)
              v1.3 — refresh_token 保管・自動更新実装, 並行リフレッシュ排他, _expireForTest(?debug=1)
              v1.4 — _expireForTest: ?debug=1 がリダイレクト後消えても sessionStorage で引き継ぐ
              v1.5 — v1.4 バグ修正: _showSplash での debug フラグ削除を廃止。削除は logout 時のみ
              v1.6 — Graph権限範囲拡大(Files.Read.All/ReadWrite, 失敗時 narrow スコープへフォールバック), SharePoint 読み取り関数(sp.getJson/list)
              v1.7 — redirectUri オプション追加(init), onLogout フック追加
              v1.8 — Graph スコープを 3 段階化 (Mail.ReadWrite 付き → WIDE → NARROW)
              v1.9 — 공통 마스터 (LEENAI.master.get/find/list, ⚙ マスター設定 패널, 충돌·입력 검사, 변경 기록 50건) */
(function(){
'use strict';

/* ── 定数 ── */
const CLIENT_ID = 'f071a165-5e9b-44bf-b6a1-baba654524db';
const TENANT_ID = 'a56988ff-c2b8-4df8-9727-889ad4205198';
const DV_URL    = 'https://orgde512c6f.crm7.dynamics.com';
const AUTH_URL  = 'https://login.microsoftonline.com/'+TENANT_ID+'/oauth2/v2.0';

/* Graph スコープ: まず広い範囲を要求、拒否されたら狭い範囲にフォールバック */
const GR_SCOPE_WIDE   = 'https://graph.microsoft.com/User.Read https://graph.microsoft.com/Mail.Send https://graph.microsoft.com/Files.Read.All https://graph.microsoft.com/Files.ReadWrite offline_access';
/* v1.8: 下書き作成 (Mail.ReadWrite) 付き。許可されていないアカウントでは拒否され、静かに WIDE へフォールバックする */
const GR_SCOPE_MAIL   = GR_SCOPE_WIDE + ' https://graph.microsoft.com/Mail.ReadWrite';
const GR_SCOPE_NARROW = 'https://graph.microsoft.com/User.Read https://graph.microsoft.com/Mail.Send offline_access';

/* SharePoint */
const SP_DRIVE_ID = 'b!ZK58_9YllU-vETQD3Xz0TbOInXbVCXJHnQ2280KJBArFLeM_18EgRbSeoO-TW__U';
const GRAPH_BASE  = 'https://graph.microsoft.com/v1.0/';

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
var _cfg = {};
var _dvToken = null, _dvExpiry = 0;
var _grToken = null, _grExpiry = 0;
var _refreshToken = null;
var _grScope = null;            // 成功した Graph スコープを記憶 (갱신 시 동일 scope 사용)
var _dvRefreshPromise = null;
var _grRefreshPromise = null;
var _user = null;
var _authStatus = {dataverse:'error', graph:'skip', graphError:''};
var _headerBtns = [];

/* ── ユーティリティ ── */
function b64url(buf){return btoa(String.fromCharCode.apply(null,new Uint8Array(buf))).replace(/\+/g,'-').replace(/\//g,'_').replace(/=/g,'');}
async function sha256(s){return b64url(await crypto.subtle.digest('SHA-256',new TextEncoder().encode(s)));}
function randomStr(n){var a='ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-._~';return Array.from(crypto.getRandomValues(new Uint8Array(n)),function(v){return a[v%a.length];}).join('');}

/* ── テーマ ── */
function _applyTheme(t){
  document.documentElement.setAttribute('data-theme',t);
  document.querySelectorAll('[data-ln="themeBtn"]').forEach(function(b){b.textContent=THEME_LABEL[t];});
  _setFavicon();
}
function _toggleTheme(){
  var cur=document.documentElement.getAttribute('data-theme')||'dark';
  var next=THEMES[(THEMES.indexOf(cur)+1)%THEMES.length];
  try{localStorage.setItem('leenai_theme',next);}catch(e){}
  _applyTheme(next);
}
function _setFavicon(){
  var acc=(getComputedStyle(document.documentElement).getPropertyValue('--accent')||'#4f8ef7').trim();
  var svg='<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 32 32"><rect width="32" height="32" rx="7" fill="#0f1117"/><text x="16" y="22" font-family="Segoe UI,Arial" font-size="15" font-weight="800" text-anchor="middle" fill="#e2e8f0">L<tspan fill="'+acc+'">A</tspan></text></svg>';
  var el=document.getElementById('leenaiFavicon');
  if(!el){el=document.createElement('link');el.rel='icon';el.id='leenaiFavicon';document.head.appendChild(el);}
  el.href='data:image/svg+xml,'+encodeURIComponent(svg);
}

/* ── ローディング ── */
var _loadEl=null, _loadMsgEl=null;
function _ensureLoading(){
  if(_loadEl) return;
  _loadEl=document.createElement('div'); _loadEl.className='ln-loading';
  _loadMsgEl=document.createElement('div'); _loadMsgEl.className='ln-loading-msg';
  var sp=document.createElement('div'); sp.className='ln-spinner';
  _loadEl.appendChild(sp); _loadEl.appendChild(_loadMsgEl);
  document.body.appendChild(_loadEl);
}
function _showLoading(msg){_ensureLoading();_loadMsgEl.textContent=msg||'';_loadEl.classList.add('ln-active');}
function _updateLoading(msg){if(_loadMsgEl)_loadMsgEl.textContent=msg;}
function _hideLoading(){if(_loadEl)_loadEl.classList.remove('ln-active');}

/* ── トースト ── */
var _toastWrap=null;
function _ensureToast(){
  if(_toastWrap) return;
  _toastWrap=document.createElement('div'); _toastWrap.className='ln-toast-wrap';
  document.body.appendChild(_toastWrap);
}
function _toast(msg, type){
  type=type||'info';
  _ensureToast();
  var el=document.createElement('div'); el.className='ln-toast ln-'+type; el.textContent=msg;
  _toastWrap.appendChild(el);
  requestAnimationFrame(function(){requestAnimationFrame(function(){el.classList.add('ln-in');});});
  setTimeout(function(){el.classList.remove('ln-in');setTimeout(function(){el.remove();},300);},3500);
}

/* ── DOM 構築 ── */
function _buildSplash(){
  var d=document.createElement('div'); d.id='ln-splash'; d.className='ln-splash';
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
    +'<div class="ln-footer">✦ Powered by LEENAI Automation System</div>';
  document.body.insertBefore(d, document.body.firstChild);
  document.getElementById('ln-login-btn').addEventListener('click', _startLogin);
}

function _buildHeader(){
  var h=document.createElement('header'); h.className='ln-header'; h.id='ln-header';
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
    +'<button class="ln-btn-ghost" onclick="LEENAI._openMasterPanel()" title="マスター設定" style="font-size:14px;padding:4px 8px;">⚙</button>'
    +'<button class="ln-btn-ghost" data-ln="themeBtn" onclick="LEENAI._toggleTheme()">🌙 ダーク</button>'
    +'<button class="ln-btn-ghost" onclick="LEENAI.logout()">ログアウト</button>'
    +'</div>';
  return h;
}

function _buildFooter(){
  var f=document.createElement('div'); f.className='ln-footer';
  f.textContent='✦ Powered by LEENAI Automation System';
  return f;
}

/* ── ヘッダードット更新 ── */
function _updateDot(){
  var dot=document.getElementById('ln-dot'); if(!dot) return;
  var s=_authStatus;
  if(s.dataverse==='ok'&&(s.graph==='ok'||s.graph==='skip')) dot.style.background='var(--ok)';
  else if(s.dataverse==='ok'&&s.graph==='error') dot.style.background='var(--warn)';
  else dot.style.background='var(--danger)';
}

/* ── ログイン (PKCE) ── */
async function _startLogin(){
  var verifier=randomStr(64);
  var challenge=await sha256(verifier);
  var state=randomStr(32);
  var pkceKey='leenai_pkce_'+_cfg.sys;
  try{sessionStorage.setItem(pkceKey,JSON.stringify({verifier:verifier,state:state}));}catch(e){}
  var rdUri=_cfg.redirectUri||(location.origin+location.pathname.replace(/index\.html$/,''));
  var scope=encodeURIComponent(DV_URL+'/user_impersonation openid profile email offline_access');
  var params='client_id='+CLIENT_ID+'&response_type=code&redirect_uri='+encodeURIComponent(rdUri)+'&scope='+scope+'&code_challenge='+challenge+'&code_challenge_method=S256&state='+state+'&prompt=select_account';
  location.href=AUTH_URL+'/authorize?'+params;
}

async function _handleCallback(){
  var url=new URL(location.href);
  var code=url.searchParams.get('code');
  var retState=url.searchParams.get('state');
  if(!code) return false;
  var pkceKey='leenai_pkce_'+_cfg.sys;
  var pkce;
  try{pkce=JSON.parse(sessionStorage.getItem(pkceKey)||'null');}catch(e){}
  sessionStorage.removeItem(pkceKey);
  if(!pkce||pkce.state!==retState){
    _toast('認証エラー: state 不一致。再ログインしてください。','danger');
    history.replaceState({},'',location.pathname);
    return false;
  }
  history.replaceState({},'',location.pathname);
  var rdUri=_cfg.redirectUri||(location.origin+location.pathname.replace(/index\.html$/,''));
  _showLoading('認証中…');
  try{
    /* ── Step1: DV トークン取得 ── */
    var r1=await fetch(AUTH_URL+'/token',{method:'POST',headers:{'Content-Type':'application/x-www-form-urlencoded'},
      body:'client_id='+CLIENT_ID+'&grant_type=authorization_code&code='+encodeURIComponent(code)+'&redirect_uri='+encodeURIComponent(rdUri)+'&code_verifier='+pkce.verifier+'&scope='+encodeURIComponent(DV_URL+'/user_impersonation openid profile email offline_access')});
    if(!r1.ok) throw new Error('DV token '+r1.status);
    var t1=await r1.json();
    _dvToken=t1.access_token; _dvExpiry=Date.now()+(t1.expires_in-60)*1000;
    _refreshToken=t1.refresh_token||null;
    _authStatus.dataverse='ok';
    /* ── id_token → user (UTF-8デコード) ── */
    try{
      var _raw=t1.id_token.split('.')[1].replace(/-/g,'+').replace(/_/g,'/');
      var _bytes=Uint8Array.from(atob(_raw),function(c){return c.charCodeAt(0);});
      var payload=JSON.parse(new TextDecoder().decode(_bytes));
      var email=payload.preferred_username||payload.email||'';
      var staff=DEFAULT_STAFF.find(function(s){return s.email.toLowerCase()===email.toLowerCase();});
      var initials=(payload.name||'').split(' ').map(function(w){return w[0];}).join('').slice(0,2).toUpperCase()||'??';
      _user={name:payload.name||email,email:email,staffCode:staff?staff.staffCode:'',initials:initials};
    }catch(e){_user={name:'(不明)',email:'',staffCode:'',initials:'??'};}
    /* ── Step2: Graph トークン取得 (wide scope → narrow フォールバック) ── */
    if(_cfg.needGraph!==false){
      _updateLoading('Outlook 接続中…');
      /* v1.8: 3 段階 — ① Mail.ReadWrite 付き → ② WIDE → ③ NARROW (②③ は従来どおり) */
      var _grTier=GR_SCOPE_MAIL;
      var r2w=await fetch(AUTH_URL+'/token',{method:'POST',headers:{'Content-Type':'application/x-www-form-urlencoded'},
        body:'client_id='+CLIENT_ID+'&grant_type=refresh_token&refresh_token='+encodeURIComponent(_refreshToken)+'&scope='+encodeURIComponent(GR_SCOPE_MAIL)});
      if(!r2w.ok){
        _grTier=GR_SCOPE_WIDE;
        r2w=await fetch(AUTH_URL+'/token',{method:'POST',headers:{'Content-Type':'application/x-www-form-urlencoded'},
          body:'client_id='+CLIENT_ID+'&grant_type=refresh_token&refresh_token='+encodeURIComponent(_refreshToken)+'&scope='+encodeURIComponent(GR_SCOPE_WIDE)});
      }
      if(r2w.ok){
        /* wide scope 成功 */
        var t2=await r2w.json();
        _grToken=t2.access_token; _grExpiry=Date.now()+(t2.expires_in-60)*1000;
        if(t2.refresh_token) _refreshToken=t2.refresh_token;
        _grScope=_grTier;
        _authStatus.graph='ok'; _authStatus.graphError='';
      } else {
        /* wide scope 失敗 → narrow スコープで再試行 */
        var ejWide=await r2w.json();
        var wideErrCode=ejWide.error||'scope_denied';
        try{
          var r2n=await fetch(AUTH_URL+'/token',{method:'POST',headers:{'Content-Type':'application/x-www-form-urlencoded'},
            body:'client_id='+CLIENT_ID+'&grant_type=refresh_token&refresh_token='+encodeURIComponent(_refreshToken)+'&scope='+encodeURIComponent(GR_SCOPE_NARROW)});
          if(!r2n.ok){var ejN=await r2n.json();throw Object.assign(new Error(r2n.status),{code:ejN.error});}
          var t2n=await r2n.json();
          _grToken=t2n.access_token; _grExpiry=Date.now()+(t2n.expires_in-60)*1000;
          if(t2n.refresh_token) _refreshToken=t2n.refresh_token;
          _grScope=GR_SCOPE_NARROW;
          _authStatus.graph='ok';
          _authStatus.graphError='files:'+wideErrCode;  // 点を主橙で警告
        }catch(e2){
          _authStatus.graph='error'; _authStatus.graphError=e2.code||String(e2);
        }
      }
    } else {_authStatus.graph='skip';}
    return true;
  }catch(e){
    _toast('ログイン失敗: '+e.message,'danger');
    _hideLoading();
    return false;
  }
}

/* ── トークン自動更新 ── */
async function _getToken(type){
  if(type==='dataverse'){
    if(_dvToken&&Date.now()<_dvExpiry) return _dvToken;
    if(!_refreshToken){
      _toast('セッションが切れました。再ログインしてください。','warn');
      setTimeout(function(){_showSplash();},2000);
      throw new Error('session expired');
    }
    if(!_dvRefreshPromise){
      _dvRefreshPromise=(async function(){
        try{
          var r=await fetch(AUTH_URL+'/token',{method:'POST',headers:{'Content-Type':'application/x-www-form-urlencoded'},
            body:'client_id='+CLIENT_ID+'&grant_type=refresh_token&refresh_token='+encodeURIComponent(_refreshToken)+'&scope='+encodeURIComponent(DV_URL+'/user_impersonation openid profile email offline_access')});
          if(!r.ok) throw new Error('DV refresh '+r.status);
          var t=await r.json();
          _dvToken=t.access_token; _dvExpiry=Date.now()+(t.expires_in-60)*1000;
          if(t.refresh_token) _refreshToken=t.refresh_token;
          _authStatus.dataverse='ok';
          return _dvToken;
        }catch(e){
          _refreshToken=null; _dvToken=null; _dvExpiry=0;
          _authStatus.dataverse='error';
          _toast('セッションが切れました。再ログインしてください。','warn');
          setTimeout(function(){_showSplash();},2000);
          throw e;
        }finally{
          _dvRefreshPromise=null;
        }
      })();
    }
    return _dvRefreshPromise;
  }
  if(type==='graph'){
    if(_authStatus.graph==='skip') return null;
    if(_grToken&&Date.now()<_grExpiry) return _grToken;
    if(!_refreshToken){
      _authStatus.graph='error'; _updateDot(); return null;
    }
    if(!_grRefreshPromise){
      _grRefreshPromise=(async function(){
        try{
          /* 成功したスコープで更新 (なければ wide を試みる) */
          var scope=_grScope||GR_SCOPE_WIDE;
          var r=await fetch(AUTH_URL+'/token',{method:'POST',headers:{'Content-Type':'application/x-www-form-urlencoded'},
            body:'client_id='+CLIENT_ID+'&grant_type=refresh_token&refresh_token='+encodeURIComponent(_refreshToken)+'&scope='+encodeURIComponent(scope)});
          if(!r.ok){var ej=await r.json();throw Object.assign(new Error(r.status),{code:ej.error});}
          var t=await r.json();
          _grToken=t.access_token; _grExpiry=Date.now()+(t.expires_in-60)*1000;
          if(t.refresh_token) _refreshToken=t.refresh_token;
          _authStatus.graph='ok';
          /* graphError は既存のまま (files: フォールバック状態を保持) */
          _updateDot();
          return _grToken;
        }catch(e){
          _authStatus.graph='error'; _authStatus.graphError=e.code||String(e);
          _updateDot();
          return null;
        }finally{
          _grRefreshPromise=null;
        }
      })();
    }
    return _grRefreshPromise;
  }
  throw new Error('unknown token type: '+type);
}

/* ── SharePoint 内部ヘルパー ── */
function _spEnc(path){
  return path.split('/').map(encodeURIComponent).join('/');
}

async function _spFetch(url, retry){
  var tok=await _getToken('graph');
  if(!tok) throw new Error('Graph token unavailable');
  var res=await fetch(url,{headers:{Authorization:'Bearer '+tok}});
  if(res.status===401&&!retry){
    /* 401: トークン強制失効 → 1回だけ再取得 */
    _grToken=null; _grExpiry=0;
    return _spFetch(url,true);
  }
  return res;
}

/* ── テスト用: 万了強制 ── */
function _isDebug(){
  try{
    if(new URL(location.href).searchParams.get('debug')==='1') return true;
    return sessionStorage.getItem('leenai_debug_'+(_cfg.sys||''))==='1';
  }catch(e){return false;}
}
function _expireForTest(){
  if(!_isDebug()) return;
  _dvExpiry=0; _grExpiry=0;
}

/* ── 画面切り替え ── */
function _showSplash(){
  var splash=document.getElementById('ln-splash');
  var app=document.getElementById('app');
  if(splash) splash.style.display='flex';
  if(app) app.style.display='none';
  _dvToken=null; _dvExpiry=0; _grToken=null; _grExpiry=0;
  _refreshToken=null; _dvRefreshPromise=null; _grRefreshPromise=null;
  _grScope=null;
  _user=null;
  _authStatus={dataverse:'error',graph:_cfg.needGraph===false?'skip':'error',graphError:''};
}

function _showApp(){
  var splash=document.getElementById('ln-splash');
  var app=document.getElementById('app');
  if(splash) splash.style.display='none';
  if(app) app.style.display='';
  if(!document.getElementById('ln-header')){
    var h=_buildHeader();
    if(app) app.insertBefore(h, app.firstChild);
  }
  if(app&&!app.querySelector('.ln-footer')) app.appendChild(_buildFooter());
  var av=document.getElementById('ln-avatar');
  var un=document.getElementById('ln-username');
  if(av&&_user) av.textContent=_user.initials||_user.staffCode||'??';
  if(un&&_user) un.textContent=_user.name;
  var extra=document.getElementById('ln-hdr-extra-btns');
  if(extra){
    extra.innerHTML='';
    _headerBtns.forEach(function(b){
      var el=document.createElement('button'); el.className='ln-btn-ghost'; el.textContent=b.label;
      if(b.title) el.title=b.title;
      el.addEventListener('click',b.onClick); extra.appendChild(el);
    });
  }
  _updateDot();
  var th='dark'; try{th=localStorage.getItem('leenai_theme')||'dark';}catch(e){}
  _applyTheme(th);
}

/* ── init ── */
async function _init(cfg){
  _cfg=Object.assign({needGraph:true},cfg);
  document.title='LEENAI '+_cfg.sys+' '+_cfg.name+' v'+_cfg.version;
  try{
    if(new URL(location.href).searchParams.get('debug')==='1')
      sessionStorage.setItem('leenai_debug_'+_cfg.sys,'1');
  }catch(e){}
  Array.from(document.body.children).forEach(function(el){if(!el.id||el.id!=='ln-splash') el.style.display='none';});
  var th='dark'; try{th=localStorage.getItem('leenai_theme')||'dark';}catch(e){}
  _applyTheme(th);
  _buildSplash();
  var url=new URL(location.href);
  if(url.searchParams.has('code')){
    var ok=await _handleCallback();
    if(ok){_hideLoading();_showApp();if(typeof _cfg.onReady==='function') _cfg.onReady();}
    else {_showSplash();}
  } else {_showSplash();}
}

/* ── logout ── */
function _logout(){
  _dvToken=null; _dvExpiry=0; _grToken=null; _grExpiry=0;
  _refreshToken=null; _dvRefreshPromise=null; _grRefreshPromise=null;
  _grScope=null;
  _user=null;
  _authStatus={dataverse:'error',graph:_cfg.needGraph===false?'skip':'error',graphError:''};
  try{sessionStorage.removeItem('leenai_debug_'+(_cfg.sys||''));}catch(e){}
  if(typeof _cfg.onLogout==='function') try{_cfg.onLogout();}catch(e){}
  _showSplash();
}

/* ── addHeaderButton ── */
function _addHeaderButton(opts){
  _headerBtns.push(opts);
  var extra=document.getElementById('ln-hdr-extra-btns');
  if(extra){
    var el=document.createElement('button'); el.className='ln-btn-ghost'; el.textContent=opts.label;
    if(opts.title) el.title=opts.title;
    el.addEventListener('click',opts.onClick); extra.insertBefore(el,extra.firstChild);
  }
}

/* ── マスター (v1.9) ── */
var _masterCache={};   // {name:{doc,eTag,webUrl}}
var _masterPanel=null, _masterCurName=null, _masterTab='table', _masterDraft={};
var MASTER_PATH='古紙/-. AI 活用/MASTER';

function _mEsc(s){return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');}

function _masterShowBadge(msg){
  var el=document.getElementById('ln-master-badge');
  if(!el){
    el=document.createElement('span'); el.id='ln-master-badge';
    el.style.cssText='font-size:11px;color:var(--warn);font-weight:600;margin-right:6px;';
    var r=document.getElementById('ln-hdr-right');
    if(r) r.insertBefore(el,r.firstChild);
  }
  el.textContent=msg||'';
}

async function _masterGet(name,opts){
  if(!(opts&&opts.fresh)&&_masterCache[name]) return _masterCache[name].doc.items;
  var contentUrl=GRAPH_BASE+'drives/'+SP_DRIVE_ID+'/root:/'+_spEnc(MASTER_PATH+'/master_'+name+'.json')+':/content';
  var metaUrl=GRAPH_BASE+'drives/'+SP_DRIVE_ID+'/root:/'+_spEnc(MASTER_PATH+'/master_'+name+'.json');
  try{
    var res=await _spFetch(contentUrl,false);
    if(!res.ok) throw new Error('SP master '+res.status);
    var eTag=res.headers.get('ETag')||res.headers.get('etag')||'';
    var doc=await res.json();
    _masterCache[name]={doc:doc,eTag:eTag,webUrl:null};
    /* フォルダの webUrl を取得 (ファイルの webUrl は JSON ダウンロードになるため) */
    try{
      var folderMeta=await _spFetch(GRAPH_BASE+'drives/'+SP_DRIVE_ID+'/root:/'+_spEnc(MASTER_PATH),false);
      if(folderMeta.ok){var fm=await folderMeta.json();_masterCache[name].webUrl=fm.webUrl||null;}
    }catch(e2){}
    try{localStorage.setItem('leenai_master_'+name,JSON.stringify(doc));}catch(e3){}
    return doc.items;
  }catch(e){
    var saved=null; try{saved=JSON.parse(localStorage.getItem('leenai_master_'+name)||'null');}catch(e4){}
    if(saved){_masterShowBadge('⚠ マスター: 前回の内容で動作中'); if(!_masterCache[name]) _masterCache[name]={doc:saved,eTag:'',webUrl:null}; return saved.items;}
    throw e;
  }
}

function _masterFind(name,filter){
  if(!_masterCache[name]) return null;
  var items=_masterCache[name].doc.items;
  return items.find(function(row){return Object.keys(filter).every(function(k){return String(row[k]||'')===String(filter[k]||'');});}) || null;
}

async function _masterList(){
  var files=await LEENAI.sp.list(MASTER_PATH);
  var masters=files.filter(function(f){return /^master_[a-z_]+\.json$/.test(f.name);});
  var results=[];
  for(var i=0;i<masters.length;i++){
    var n=masters[i].name.replace(/^master_/,'').replace(/\.json$/,'');
    try{
      await _masterGet(n,{fresh:false});
      var d=_masterCache[n].doc;
      results.push({name:n,title:d.title||n,count:(d.items||[]).length,updated_at:d.updated_at||'',updated_by:d.updated_by||''});
    }catch(e){results.push({name:n,title:n,count:0,updated_at:'',updated_by:''});}
  }
  return results;
}

function _buildMasterPanel(){
  if(_masterPanel) return _masterPanel;
  var panel=document.createElement('div');
  panel.className='ln-master-panel ln-hidden'; panel.id='ln-master-panel';
  panel.innerHTML=
    '<div class="ln-master-topbar">'
    +'<span class="ln-logo">LEEN<span>AI</span></span>'
    +'<h2>⚙ マスター設定</h2>'
    +'<span id="ln-master-readonly-msg" style="font-size:11px;color:var(--warn);display:none;">編集権限がありません（読取専用）</span>'
    +'<button class="ln-btn-ghost" id="ln-master-close">✕ 閉じる</button>'
    +'</div>'
    +'<div class="ln-master-body">'
    +'<div class="ln-master-sidebar" id="ln-master-sidebar"><div style="padding:12px;font-size:12px;color:var(--text3);">読み込み中…</div></div>'
    +'<div class="ln-master-main" id="ln-master-main">'
    +'<div class="ln-master-tabs">'
    +'<button class="ln-master-tab ln-active" id="ln-mtab-table" onclick="LEENAI._masterTabSwitch(\'table\')">📋 データ</button>'
    +'<button class="ln-master-tab" id="ln-mtab-history" onclick="LEENAI._masterTabSwitch(\'history\')">📜 履歴</button>'
    +'</div>'
    +'<div id="ln-master-content" style="display:flex;flex-direction:column;flex:1;overflow:hidden;"></div>'
    +'</div></div>';
  document.body.appendChild(panel);
  document.getElementById('ln-master-close').addEventListener('click',function(){_closeMasterPanel();});
  _masterPanel=panel;
  return panel;
}

function _openMasterPanel(){
  _buildMasterPanel();
  _masterPanel.classList.remove('ln-hidden');
  _loadMasterSidebar();
}

function _closeMasterPanel(){
  if(_masterPanel) _masterPanel.classList.add('ln-hidden');
}

async function _loadMasterSidebar(){
  var sb=document.getElementById('ln-master-sidebar');
  if(!sb) return;
  sb.innerHTML='<div style="padding:12px;font-size:12px;color:var(--text3);">読み込み中…</div>';
  try{
    var list=await _masterList();
    sb.innerHTML='';
    list.forEach(function(m){
      var d=document.createElement('div');
      d.className='ln-master-sitem'+(m.name===_masterCurName?' ln-active':'');
      d.setAttribute('data-mname',m.name);
      d.innerHTML='<div style="font-weight:600;">'+_mEsc(m.title)+'</div>'
        +'<div style="font-size:10px;color:var(--text3);">'+m.count+'件</div>';
      d.addEventListener('click',function(){_selectMaster(m.name);});
      sb.appendChild(d);
    });
  }catch(e){
    sb.innerHTML='<div style="padding:12px;font-size:11px;color:var(--danger);">読込エラー: '+_mEsc(e.message)+'</div>';
  }
}

async function _selectMaster(name){
  _masterCurName=name;
  document.querySelectorAll('#ln-master-sidebar .ln-master-sitem').forEach(function(el){
    el.classList.toggle('ln-active',el.getAttribute('data-mname')===name);
  });
  _masterTab='table';
  var t1=document.getElementById('ln-mtab-table'); if(t1) t1.classList.add('ln-active');
  var t2=document.getElementById('ln-mtab-history'); if(t2) t2.classList.remove('ln-active');
  await _renderMasterTable(name);
}

function _masterIsReadonly(){
  return !!(_authStatus.graphError&&_authStatus.graphError.startsWith('files:'));
}

async function _renderMasterTable(name){
  var content=document.getElementById('ln-master-content');
  if(!content) return;
  content.innerHTML='<div style="padding:12px;font-size:12px;color:var(--text3);">読み込み中…</div>';
  try{
    if(!_masterCache[name]) await _masterGet(name,{fresh:true});
    var doc=_masterCache[name].doc;
    var fields=doc.fields;
    if(!_masterDraft[name]) _masterDraft[name]={orig:JSON.parse(JSON.stringify(doc.items)),rows:JSON.parse(JSON.stringify(doc.items)),deleted:[],added:[]};
    var draft=_masterDraft[name];
    var ro=_masterIsReadonly();
    var rmEl=document.getElementById('ln-master-readonly-msg');
    if(rmEl) rmEl.style.display=ro?'':'none';
    var html='<div class="ln-master-toolbar">'
      +'<input class="ln-master-search" id="ln-msearch" placeholder="検索…" oninput="LEENAI._masterSearch()">'
      +(!ro?'<button class="ln-btn-ghost" style="font-size:12px;" onclick="LEENAI._masterAddRow()">＋ 行を追加</button>':'')
      +'</div>'
      +'<div class="ln-master-table-wrap">'
      +'<table class="ln-master-table" id="ln-master-tbl"><thead><tr>';
    fields.forEach(function(f){html+='<th onclick="LEENAI._masterSort(\''+f.key+'\')" style="cursor:pointer;">'+_mEsc(f.label)+' ⇅</th>';});
    html+='<th>操作</th></tr></thead><tbody id="ln-master-tbody"></tbody></table></div>'
      +'<div class="ln-master-footer">'
      +(!ro?'<button class="ln-btn-ghost" style="font-size:12px;background:var(--accent-btn);color:#fff;border:none;" onclick="LEENAI._masterSave()">💾 保存</button>':'')
      +(!ro?'<button class="ln-btn-ghost" style="font-size:12px;" onclick="LEENAI._masterRevert()">↩ 元に戻す</button>':'')
      +'<a href="#" id="ln-master-splink" style="font-size:11px;color:var(--accent);" target="_blank">SharePoint で開く</a>'
      +'<button class="ln-btn-ghost" style="font-size:12px;" onclick="LEENAI._masterDownload()">⬇ JSON ダウンロード</button>'
      +'<span class="ln-master-info" id="ln-master-info"></span>'
      +'</div>'
      +'<div class="ln-master-err" id="ln-master-err"></div>';
    content.innerHTML=html;
    _masterRenderRows(name,fields,draft.rows,ro);
    var linkEl=document.getElementById('ln-master-splink');
    if(linkEl&&_masterCache[name]&&_masterCache[name].webUrl) linkEl.href=_masterCache[name].webUrl;
  }catch(e){
    content.innerHTML='<div style="padding:20px;color:var(--danger);">エラー: '+_mEsc(e.message)+'</div>';
  }
}

function _masterRenderRows(name,fields,rows,ro){
  var tbody=document.getElementById('ln-master-tbody'); if(!tbody) return;
  var search=''; try{search=(document.getElementById('ln-msearch')||{}).value||'';}catch(e){}
  var draft=_masterDraft[name]; if(!draft) return;
  tbody.innerHTML='';
  rows.forEach(function(row,ri){
    if(search){
      var match=fields.some(function(f){return String(row[f.key]||'').toLowerCase().includes(search.toLowerCase());});
      if(!match) return;
    }
    var isNew=draft.added.indexOf(row)>=0;
    var isDel=draft.deleted.indexOf(ri)>=0;
    var isChanged=!isNew&&JSON.stringify(row)!==JSON.stringify(draft.orig[ri]);
    var tr=document.createElement('tr');
    tr.className=isDel?'ln-master-row-del':isNew?'ln-master-row-new':isChanged?'ln-master-row-changed':'';
    var cells='';
    fields.forEach(function(f){
      var v=row[f.key]!=null?String(row[f.key]):'';
      if(ro||isDel){cells+='<td>'+_mEsc(v)+'</td>';}
      else if(f.type==='bool'){cells+='<td><input type="checkbox"'+(v==='true'||v===true?' checked':'')
        +' data-ri="'+ri+'" data-key="'+f.key+'" onchange="LEENAI._masterCellChange(this)"></td>';}
      else{cells+='<td><input class="ln-master-cell-input" type="text" value="'+_mEsc(v)+'" data-ri="'+ri+'" data-key="'+f.key+'" oninput="LEENAI._masterCellChange(this)"></td>';}
    });
    var act=ro?'<td></td>':isDel
      ?'<td><button onclick="LEENAI._masterRestoreRow('+ri+')">↩</button></td>'
      :'<td><button onclick="LEENAI._masterDeleteRow('+ri+')">🗑</button></td>';
    tr.innerHTML=cells+act;
    tbody.appendChild(tr);
  });
  var info=document.getElementById('ln-master-info'); if(info) info.textContent=rows.length+'件';
}

async function _renderMasterHistory(name){
  var content=document.getElementById('ln-master-content'); if(!content) return;
  try{
    if(!_masterCache[name]) await _masterGet(name,{fresh:true});
    var history=_masterCache[name].doc.history||[];
    var html='<div style="overflow-y:auto;padding:8px;">';
    if(!history.length) html+='<div style="padding:20px;color:var(--text3);">履歴なし</div>';
    history.forEach(function(h){
      html+='<div class="ln-master-history-item"><div>'+_mEsc(h.summary||'')+'</div>'
        +'<div class="ln-mh-meta">'+_mEsc(h.by||'')+'&nbsp;·&nbsp;'+_mEsc(h.at||'')+'</div></div>';
    });
    html+='</div>';
    content.innerHTML=html;
  }catch(e){content.innerHTML='<div style="padding:20px;color:var(--danger);">エラー: '+_mEsc(e.message)+'</div>';}
}

function _masterTabSwitch(tab){
  _masterTab=tab;
  var t1=document.getElementById('ln-mtab-table'); if(t1) t1.classList.toggle('ln-active',tab==='table');
  var t2=document.getElementById('ln-mtab-history'); if(t2) t2.classList.toggle('ln-active',tab==='history');
  if(tab==='table') _renderMasterTable(_masterCurName);
  else _renderMasterHistory(_masterCurName);
}

function _masterSearch(){
  if(!_masterCurName||!_masterCache[_masterCurName]) return;
  var d=_masterCache[_masterCurName];
  _masterRenderRows(_masterCurName,d.doc.fields,_masterDraft[_masterCurName]?_masterDraft[_masterCurName].rows:d.doc.items,_masterIsReadonly());
}

function _masterSort(key){
  if(!_masterCurName||!_masterDraft[_masterCurName]) return;
  _masterDraft[_masterCurName].rows.sort(function(a,b){return String(a[key]||'').localeCompare(String(b[key]||''),'ja');});
  var d=_masterCache[_masterCurName];
  _masterRenderRows(_masterCurName,d.doc.fields,_masterDraft[_masterCurName].rows,_masterIsReadonly());
}

function _masterAddRow(){
  if(!_masterCurName||!_masterDraft[_masterCurName]) return;
  var d=_masterCache[_masterCurName].doc; var row={};
  d.fields.forEach(function(f){row[f.key]='';});
  _masterDraft[_masterCurName].rows.push(row);
  _masterDraft[_masterCurName].added.push(row);
  _masterRenderRows(_masterCurName,d.fields,_masterDraft[_masterCurName].rows,false);
}

function _masterDeleteRow(ri){
  if(!confirm('この行を削除します。よろしいですか？')) return;
  _masterDraft[_masterCurName].deleted.push(ri);
  var d=_masterCache[_masterCurName].doc;
  _masterRenderRows(_masterCurName,d.fields,_masterDraft[_masterCurName].rows,false);
}

function _masterRestoreRow(ri){
  var di=_masterDraft[_masterCurName].deleted.indexOf(ri);
  if(di>=0) _masterDraft[_masterCurName].deleted.splice(di,1);
  var d=_masterCache[_masterCurName].doc;
  _masterRenderRows(_masterCurName,d.fields,_masterDraft[_masterCurName].rows,false);
}

function _masterCellChange(el){
  var ri=parseInt(el.dataset.ri),key=el.dataset.key;
  if(_masterDraft[_masterCurName]) _masterDraft[_masterCurName].rows[ri][key]=el.type==='checkbox'?el.checked:el.value;
}

function _masterRevert(){
  delete _masterDraft[_masterCurName];
  _renderMasterTable(_masterCurName);
}

function _masterDownload(){
  if(!_masterCurName||!_masterCache[_masterCurName]) return;
  var blob=new Blob([JSON.stringify(_masterCache[_masterCurName].doc,null,2)],{type:'application/json'});
  var a=document.createElement('a'); a.href=URL.createObjectURL(blob);
  a.download='master_'+_masterCurName+'.json'; a.click();
}

async function _masterSave(){
  if(!_masterCurName) return;
  var draft=_masterDraft[_masterCurName]; var cached=_masterCache[_masterCurName];
  if(!draft||!cached) return;
  var doc=JSON.parse(JSON.stringify(cached.doc));
  var fields=doc.fields;
  var errEl=document.getElementById('ln-master-err'); if(errEl) errEl.textContent='';
  var rows=draft.rows.filter(function(_,i){return draft.deleted.indexOf(i)<0;});
  var errors=[];
  rows.forEach(function(row){
    fields.forEach(function(f){
      var v=String(row[f.key]||'').trim();
      if(f.required&&!v){errors.push(f.label+': 必須項目です');}
      if(v&&f.type==='email'){if(!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(v)) errors.push(f.label+': メール形式エラー ('+v+')');}
      if(v&&f.type==='emails'){v.split(/[;,\/\s]+/).filter(Boolean).forEach(function(a){if(!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(a.trim())) errors.push(f.label+': メール形式 ('+a.trim()+')');});}
      if(v&&f.type==='emails_or_fax'){v.split(/[;,\/\s]+/).filter(Boolean).forEach(function(a){if(a.trim()!=='FAX'&&!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(a.trim())) errors.push(f.label+': メールかFAX ('+a.trim()+')');});}
    });
  });
  if(doc.unique&&doc.unique.length){
    var seen={};
    rows.forEach(function(row){var k=doc.unique.map(function(u){return row[u]||'';}).join('|'); if(seen[k]) errors.push('キー重複: '+k); seen[k]=true;});
  }
  if(errors.length){if(errEl) errEl.textContent=errors.slice(0,3).join(' / '); return;}
  var addedCt=draft.added.length;
  var deletedCt=draft.deleted.length;
  var changedCt=rows.filter(function(row,i){return draft.added.indexOf(row)<0&&JSON.stringify(row)!==JSON.stringify(draft.orig[i]);}).length;
  var summary='追加 '+addedCt+'件・修正 '+changedCt+'件・削除 '+deletedCt+'件';
  doc.items=rows;
  doc.updated_at=new Date().toLocaleString('sv',{timeZone:'Asia/Tokyo'}).replace(' ','T')+'+09:00';
  doc.updated_by=(_user&&_user.name)||'?';
  if(!doc.history) doc.history=[];
  doc.history.unshift({at:doc.updated_at,by:doc.updated_by,summary:summary});
  if(doc.history.length>50) doc.history=doc.history.slice(0,50);
  var body=JSON.stringify(doc,null,2);
  var putUrl=GRAPH_BASE+'drives/'+SP_DRIVE_ID+'/root:/'+_spEnc(MASTER_PATH+'/master_'+_masterCurName+'.json')+':/content';
  try{
    var tok=await _getToken('graph');
    if(!tok) throw new Error('Graph トークンがありません');
    var hdrs={Authorization:'Bearer '+tok,'Content-Type':'application/json'};
    if(cached.eTag) hdrs['If-Match']=cached.eTag;
    var res=await fetch(putUrl,{method:'PUT',headers:hdrs,body:body});
    if(res.status===412){
      if(errEl){
        errEl.textContent='他の人が更新しました。再読み込みしてください。';
        var rb=document.createElement('button'); rb.className='ln-btn-ghost'; rb.style.marginLeft='8px'; rb.style.fontSize='12px';
        rb.textContent='再読み込み';
        rb.onclick=function(){delete _masterDraft[_masterCurName]; delete _masterCache[_masterCurName]; _renderMasterTable(_masterCurName);};
        errEl.appendChild(rb);
      }
      return;
    }
    if(!res.ok) throw new Error('PUT '+res.status);
    var saved=await res.json();
    cached.doc=doc; cached.eTag=saved['@odata.etag']||res.headers.get('ETag')||cached.eTag;
    try{localStorage.setItem('leenai_master_'+_masterCurName,JSON.stringify(doc));}catch(e2){}
    delete _masterDraft[_masterCurName];
    _toast('保存しました ('+summary+')','ok');
    _renderMasterTable(_masterCurName);
    _loadMasterSidebar();
  }catch(e){if(errEl) errEl.textContent='保存エラー: '+e.message;}
}

/* ── 公開 API ── */
window.LEENAI={
  init:_init,
  auth:{
    get status(){return Object.assign({},_authStatus);},
    token:_getToken,
    _expireForTest:_expireForTest,
  },
  get user(){return _user?Object.assign({},_user):null;},
  logout:_logout,
  loading:{show:_showLoading,update:_updateLoading,hide:_hideLoading},
  toast:_toast,
  addHeaderButton:_addHeaderButton,
  STAFF:DEFAULT_STAFF,
  COMPANY:{tel:'+81-3-3528-9850',fax:'+81-3-3528-9851'},
  VERSION:'v1.9',
  _toggleTheme:_toggleTheme,
  _openMasterPanel:_openMasterPanel,
  _masterTabSwitch:_masterTabSwitch,
  _masterSearch:_masterSearch,
  _masterSort:_masterSort,
  _masterAddRow:_masterAddRow,
  _masterDeleteRow:_masterDeleteRow,
  _masterRestoreRow:_masterRestoreRow,
  _masterCellChange:_masterCellChange,
  _masterRevert:_masterRevert,
  _masterDownload:_masterDownload,
  _masterSave:_masterSave,
  /* ── マスター (v1.9) ── */
  master:{
    get:_masterGet,
    find:_masterFind,
    list:_masterList,
  },
  /* ── SharePoint 読み取り (v1.6) ── */
  sp:{
    getJson:async function(path){
      var url=GRAPH_BASE+'drives/'+SP_DRIVE_ID+'/root:/'+_spEnc(path)+':/content';
      var res=await _spFetch(url,false);
      if(res.status===404) return null;
      if(!res.ok) throw new Error('SP getJson '+res.status);
      return res.json();
    },
    list:async function(folderPath){
      var results=[];
      var url=GRAPH_BASE+'drives/'+SP_DRIVE_ID+'/root:/'+_spEnc(folderPath)+':/children?$select=name,lastModifiedDateTime,size&$top=100';
      while(url){
        var res=await _spFetch(url,false);
        if(!res.ok) throw new Error('SP list '+res.status);
        var data=await res.json();
        (data.value||[]).forEach(function(f){
          if(f.name) results.push({name:f.name,lastModified:f.lastModifiedDateTime,size:f.size});
        });
        url=data['@odata.nextLink']||null;
      }
      return results;
    },
  },
  SP:{
    DRIVE_ID:SP_DRIVE_ID,
    PATH:{
      LC_JSON:[
        '古紙/-. AI 活用/D-01-LC_FD_CHECKER',
        '古紙/-. 古紙書類/▲ LC_LEENAI_JSON',
      ],
      MASTER:MASTER_PATH,
    },
  },
};
})();
