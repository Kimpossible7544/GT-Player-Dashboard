// Shared access gate for every GT dashboard page. Include in <head>; the page
// body is hidden until the alliance code is entered (or already stored for the
// session under the same key the Goals / Since Aug 31 pages use).
(function(){
  var ACCESS_CODE = 'fucknabs';
  var AUTH_KEY = 'gtGoalsAuth';
  var subtitle = (document.currentScript && document.currentScript.getAttribute('data-title')) || document.title;

  function isAuthed(){ return sessionStorage.getItem(AUTH_KEY) === '1'; }

  var style = document.createElement('style');
  style.textContent =
    'html.gt-locked body>:not(#gtLoginWrap){display:none!important;}' +
    '#gtLoginWrap{position:fixed;top:0;left:0;right:0;bottom:0;background:var(--bg,#0a0a0f);display:flex;align-items:center;justify-content:center;z-index:9999;}' +
    '#gtLoginWrap .login-box{background:var(--panel,#111118);border:1px solid var(--border,#2a2a3a);border-radius:16px;padding:40px 36px;width:100%;max-width:400px;text-align:center;font-family:"Segoe UI",system-ui,sans-serif;}' +
    '#gtLoginWrap .login-logo{font-size:1.4rem;font-weight:900;color:var(--accent,#00ffb4);letter-spacing:3px;margin-bottom:6px;}' +
    '#gtLoginWrap .login-sub{font-size:0.8rem;color:var(--dim,#6b6b88);letter-spacing:1px;margin-bottom:28px;}' +
    '#gtLoginWrap .login-input{width:100%;background:var(--panel2,#16161f);border:1px solid var(--border,#2a2a3a);color:var(--text,#dcdce8);padding:12px 16px;border-radius:8px;font-size:1.1rem;font-family:"Courier New",monospace;outline:none;text-align:center;letter-spacing:3px;margin-bottom:12px;transition:border-color 0.2s;box-sizing:border-box;}' +
    '#gtLoginWrap .login-input:focus{border-color:var(--accent,#00ffb4);}' +
    '#gtLoginWrap .login-btn{width:100%;background:var(--accent,#00ffb4);color:var(--bg,#0a0a0f);border:none;padding:12px;border-radius:8px;font-size:0.9rem;font-weight:900;letter-spacing:2px;text-transform:uppercase;cursor:pointer;transition:opacity 0.2s;}' +
    '#gtLoginWrap .login-btn:hover{opacity:0.85;}' +
    '#gtLoginWrap .login-error{color:var(--red,#ff4466);font-size:0.8rem;margin-top:10px;min-height:18px;}' +
    '@media(max-width:768px){#gtLoginWrap .login-box{padding:32px 24px;margin:20px;}}';
  document.head.appendChild(style);

  function showMemberLinks(){
    ['goalsNavLink','recentNavLink'].forEach(function(id){
      var el = document.getElementById(id);
      if (el) el.style.display = '';
    });
  }

  function unlock(){
    sessionStorage.setItem(AUTH_KEY, '1');
    document.documentElement.classList.remove('gt-locked');
    var wrap = document.getElementById('gtLoginWrap');
    if (wrap) wrap.parentNode.removeChild(wrap);
    showMemberLinks();
    window.dispatchEvent(new Event('resize')); // charts hidden during load re-measure
  }

  function buildLogin(){
    var wrap = document.createElement('div');
    wrap.id = 'gtLoginWrap';
    wrap.innerHTML =
      '<div class="login-box">' +
        '<div class="login-logo">⚔ GT ALLIANCE</div>' +
        '<div class="login-sub"></div>' +
        '<input class="login-input" id="gtLoginInput" type="password" placeholder="Enter code" maxlength="20" autocomplete="off"/>' +
        '<button class="login-btn" id="gtLoginBtn">Enter</button>' +
        '<div class="login-error" id="gtLoginError"></div>' +
      '</div>';
    wrap.querySelector('.login-sub').textContent = subtitle;
    document.body.appendChild(wrap);
    var inp = document.getElementById('gtLoginInput');
    var err = document.getElementById('gtLoginError');
    function tryCode(){
      err.textContent = '';
      if (inp.value.trim() === ACCESS_CODE){ unlock(); return; }
      err.textContent = 'Invalid code. Please try again.';
      inp.value = '';
      inp.focus();
    }
    inp.addEventListener('keydown', function(e){ if (e.key === 'Enter') tryCode(); });
    document.getElementById('gtLoginBtn').addEventListener('click', tryCode);
    inp.focus();
  }

  if (isAuthed()){
    document.addEventListener('DOMContentLoaded', showMemberLinks);
    return;
  }
  document.documentElement.classList.add('gt-locked');
  document.addEventListener('DOMContentLoaded', buildLogin);
})();
