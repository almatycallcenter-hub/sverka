/* «ТТ Алматы» — нижние вкладки мини-приложения: Меню · Списки · Фото · Кухня · Витрина.
   Подключается на всех страницах, но работает только внутри Телеграма: в обычном
   браузере страница выглядит как раньше. Какие вкладки показать, решает профиль
   из приёмника (tg_me): в каких группах человек состоит. Пока профиль не пришёл,
   показываем то, что знаем по ссылке. */
(function(){
  'use strict';
  var W = null;
  try{ W = window.Telegram && window.Telegram.WebApp && window.Telegram.WebApp.initData ? window.Telegram.WebApp : null; }catch(e){}
  if (!W) return;
  var API = 'https://script.google.com/macros/s/AKfycbyB2zMJldSSbflPVmWqlJxWvBcRrI1agvAnGSYHeUVdvVElOTxbs3dPszB_IdqOm88zyg/exec';
  var KEY = 'tt_prof', P = {};
  try{ P = JSON.parse(sessionStorage.getItem(KEY) || '{}') || {}; }catch(e){ P = {}; }
  function save(){ try{ sessionStorage.setItem(KEY, JSON.stringify(P)); }catch(e){} }

  var qp = new URLSearchParams(location.search);
  var sp = (W.initDataUnsafe || {}).start_param || '';
  var code = qp.get('c') || (/^(ph|ls|kx)[0-9a-f]{10}$/.test(sp) ? sp.slice(2) : '') || P.code || '';
  if (code && code !== P.code){ P.code = code; save(); }
  var page = (location.pathname.split('/').pop() || 'app.html').replace(/\?.*$/, '');
  var HERE = { 'app.html': 'home', 'index.html': 'ls', '': 'ls', 'spisanie.html': 'ph',
               'kuhnya.html': 'kx', 'vitrina.html': 'vit', 'bron.html': 'bk' }[page] || 'home';
  /* «ТТ Алматы» без кода филиала на странице списков — это обычный сайт, не трогаем */
  if (HERE === 'ls' && !document.documentElement.classList.contains('tgm') &&
      !/^ls/.test(sp) && !qp.get('c')) return;

  var I = {
    home: '<path d="M4 4h6v6H4zM14 4h6v6h-6zM4 14h6v6H4zM14 14h6v6h-6z"/>',
    ls:   '<path d="M8 6h12M8 12h12M8 18h12"/><circle cx="4" cy="6" r="1"/><circle cx="4" cy="12" r="1"/><circle cx="4" cy="18" r="1"/>',
    ph:   '<path d="M4 8h3l2-3h6l2 3h3v11H4z"/><circle cx="12" cy="13" r="3.5"/>',
    kx:   '<path d="M7 14a4 4 0 0 1-.6-7.95A5.5 5.5 0 0 1 17.6 6.05 4 4 0 0 1 17 14"/><path d="M7 12v8h10v-8"/><path d="M7 17h10"/>',
    vit:  '<path d="M4 9l1.5-5h13L20 9"/><path d="M4 9h16v11H4z"/><path d="M9 20v-6h6v6"/>'
  };
  function icon(n){ return '<svg viewBox="0 0 24 24" width="24" height="24" fill="none" stroke="currentColor" ' +
    'stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">' + I[n] + '</svg>'; }

  function inBranch(c){ return !P.loaded || (P.branches || []).some(function(b){ return b.code === c; }); }
  function isKitchen(c){ return (P.kitchens || []).some(function(k){ return k.code === c; }); }
  function tabs(){
    var t = [{ id: 'home', label: 'Меню', href: 'app.html' }];
    var c = P.code || '';
    if (c && inBranch(c)){
      t.push({ id: 'ls', label: 'Списки', href: 'index.html?c=' + c });
      t.push({ id: 'ph', label: 'Фото', href: 'spisanie.html?c=' + c });
    }
    var kc = c && isKitchen(c) ? c : (P.baker && P.kitchens && P.kitchens[0] ? (P.kx || P.kitchens[0].code) : '');
    if (HERE === 'kx' && !kc) kc = c;
    if (kc) t.push({ id: 'kx', label: 'Кухня', href: 'kuhnya.html?c=' + kc });
    var k = P.show || '';
    try{ k = k || localStorage.getItem('tt_vitrina_key') || ''; }catch(e){}
    t.push({ id: 'vit', label: 'Витрина', href: 'vitrina.html' + (k ? '?k=' + encodeURIComponent(k) : '?tg=1') });
    return t;
  }

  var css = document.createElement('style');
  css.textContent =
    ':root{--ttbar:calc(62px + env(safe-area-inset-bottom))}' +
    'html.ttbar body{padding-bottom:calc(var(--ttbar) + 104px)!important}' +
    'html.ttbar .bar,html.ttbar.tgm .page.on .tgBar{bottom:var(--ttbar)!important;' +
      'padding-bottom:10px!important;z-index:30}' +
    'html.ttbar .toast{bottom:calc(var(--ttbar) + 86px)!important}' +
    'html.ttbar .sheet{z-index:60!important}' +
    '.ttnav{position:fixed;left:0;right:0;bottom:0;z-index:40;background:rgba(255,255,255,.96);' +
      '-webkit-backdrop-filter:saturate(1.4) blur(12px);backdrop-filter:saturate(1.4) blur(12px);' +
      'border-top:1px solid #e6e3dd;border-radius:22px 22px 0 0;box-shadow:0 -6px 24px rgba(20,17,12,.06);' +
      'padding:6px 6px env(safe-area-inset-bottom);display:flex;justify-content:space-around}' +
    '.ttnav a{flex:1;display:flex;flex-direction:column;align-items:center;gap:3px;padding:6px 2px 8px;' +
      'color:#8a8478;text-decoration:none;font:600 11.5px/1.1 system-ui,-apple-system,"Segoe UI",Roboto,sans-serif;' +
      'position:relative;-webkit-tap-highlight-color:transparent}' +
    '.ttnav a.on{color:#1f2d4d}' +
    '.ttnav a.on::before{content:"";position:absolute;top:-6px;left:28%;right:28%;height:3px;border-radius:0 0 3px 3px;background:#1f2d4d}' +
    '@media (prefers-color-scheme:dark){html.ttbarDark .ttnav{background:rgba(28,28,30,.96);border-color:#333}}';
  document.head.appendChild(css);

  function draw(){
    var bar = document.querySelector('nav.ttnav');
    if (!bar){ bar = document.createElement('nav'); bar.className = 'ttnav'; document.body.appendChild(bar); }
    bar.innerHTML = tabs().map(function(t){
      return '<a href="#" data-href="' + t.href + '" class="' + (t.id === HERE ? 'on' : '') + '">' +
             icon(t.id) + '<span>' + t.label + '</span></a>';
    }).join('');
    document.documentElement.classList.add('ttbar');
  }
  function go(href){
    try{ W.HapticFeedback && W.HapticFeedback.selectionChanged(); }catch(e){}
    location.replace(href + location.hash);
  }
  function start(){
    draw();
    document.querySelector('nav.ttnav').addEventListener('click', function(e){
      var a = e.target.closest('a[data-href]'); if (!a) return;
      e.preventDefault();
      if (a.classList.contains('on')) return;
      go(a.getAttribute('data-href'));
    });
    if (!P.loaded) load();
  }
  /* Профиль: в каких группах человек, какие кухни, ключ витрины */
  function load(){
    fetch(API, { method: 'POST', body: JSON.stringify({ action: 'tg_me', init: W.initData }) })
      .then(function(r){ return r.json(); })
      .then(function(r){
        if (!r || r.error) return;
        P.loaded = 1; P.name = r.name; P.branches = r.branches || []; P.kitchens = r.kitchens || [];
        P.baker = !!r.baker; P.staff = !!r.staff; P.show = r.show || '';
        if (P.code && P.branches.length && !P.branches.some(function(b){ return b.code === P.code; }) &&
            !P.kitchens.some(function(k){ return k.code === P.code; })) P.code = '';
        if (!P.code && P.branches.length) P.code = P.branches[0].code;
        save(); draw();
        try{ window.dispatchEvent(new CustomEvent('ttprofile', { detail: P })); }catch(e){}
      }).catch(function(){});
  }
  window.TTBAR = { profile: function(){ return P; }, go: go, reload: function(){ P.loaded = 0; save(); load(); },
                   setCode: function(c){ P.code = c; save(); draw(); },
                   setKx: function(c){ P.kx = c; save(); draw(); } };
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', start); else start();
})();
