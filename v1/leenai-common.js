/* LEENAI_COMMON v1.7 (2026-10-08) — DOC-Z-05 v1.0
   window.LEENAI を 1 つだけ公開。他のグローバル変数は作らない。
   ChangeLog: v1.0 — 新規
              v1.1 — id_token UTF-8デコード修正(TextDecoder), staffCode をメール prefix 基準に変更, 未登録者は空文字
              v1.2 — バージョン番号修正 (v1.1 → v1.2)
              v1.3 — refresh_token 保管・自動更新実装, 並行リフレッシュ排他, _expireForTest(?debug=1)
              v1.4 — _expireForTest: ?debug=1 がリダイレクト後消えても sessionStorage で引き継ぐ
              v1.5 — v1.4 バグ修正: _showSplash での debug フラグ削除を廃止。削除は logout 時のみ
              v1.6 — Graph権限範囲拡大(Files.Read.All/ReadWrite, 失敗時 narrow スコープへフォールバック), SharePoint 読み取り関数(sp.getJson/list)
              v1.7 — redirectUri オプション追加(init), onLogout フック追加 */
(function(){
'use strict';

/* ── 定数 ── */
const CLIENT_ID = 'f071a165-5e9b-44bf-b6a1-baba654524db';
const TENANT_ID = 'a56988ff-c2b8-4df8-9727-889ad4205198';
const DV_URL    = 'https://orgde512c6f.crm7.dynamics.com';
const AUTH_URL  = 'https://login.microsoftonline.com/'+TENANT_ID+'/oauth2/v2.0';

/* Graph スコープ: まず広い範囲を要求、拒否されたら狭い範囲にフォールバック */
const GR_SCOPE_WIDE   = 'https://graph.microsoft.com/User.Read https://graph.microsoft.com/Mail.Send https://graph.microsoft.com/Files.Read.All https://graph.microsoft.com/Files.ReadWrite offline_access';
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
      var r2w=await fetch(AUTH_URL+'/token',{method:'POST',headers:{'Content-Type':'application/x-www-form-urlencoded'},
        body:'client_id='+CLIENT_ID+'&grant_type=refresh_token&refresh_token='+encodeURIComponent(_refreshToken)+'&scope='+encodeURIComponent(GR_SCOPE_WIDE)});
      if(r2w.ok){
        /* wide scope 成功 */
        var t2=await r2w.json();
        _grToken=t2.access_token; _grExpiry=Date.now()+(t2.expires_in-60)*1000;
        if(t2.refresh_token) _refreshToken=t2.refresh_token;
        _grScope=GR_SCOPE_WIDE;
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
  VERSION:'v1.7',
  _toggleTheme:_toggleTheme,
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
    },
  },
};
})();
