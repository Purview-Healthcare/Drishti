/* Inside Brahmastra: ?embed=1&theme=light|navy|gold
   The SLA tab opens Drishti with these parameters. The page then hides its own branding (the nav logo, the welcome
   card's logo and byline, the Feedback tab) and takes Brahmastra's colours, type and flat cards, in whichever of
   Brahmastra's themes the person picked, so it reads as one app. Opened directly, without ?embed=1, nothing here runs
   and Drishti looks exactly as before. Numbers, data and controls are untouched either way. */
(function () {
  'use strict';
  var q; try { q = new URLSearchParams(location.search); } catch (e) { return; }
  if (q.get('embed') !== '1') return;
  var bt = { light: 1, navy: 1, gold: 1 }[q.get('theme')] ? q.get('theme') : 'light';
  var root = document.documentElement;
  root.classList.add('dr-embed'); root.setAttribute('data-bt', bt);
  // Brahmastra's three themes: the page, its panels and lines, the lettering, the accent, its primary button and the
  // colours of the overview's hero band
  var T = {
    light: { bg: '#EEF1F4', panel: '#FFFFFF', border: '#DFE4EA', soft: '#EDF1F4', text: '#0F1B24', dim: '#3B4A55', faint: '#627079', accent: '#0E8C74', accent2: '#12B886',
      pri: 'linear-gradient(120deg,#22CAE4 0%,#2FD1B4 48%,#30E494 100%)', priInk: '#06231C', hero: 'linear-gradient(135deg,#0B4F46,#0E8C74)' },
    navy: { bg: '#071E33', panel: '#0B2740', border: 'rgba(255,255,255,.12)', soft: 'rgba(255,255,255,.07)', text: '#FFFFFF', dim: 'rgba(255,255,255,.82)', faint: 'rgba(255,255,255,.60)', accent: '#2BD4A4', accent2: '#38C8E8',
      pri: '#2FD9B3', priInk: '#06231C', hero: 'linear-gradient(135deg,#0B2740,#143F63)' },
    gold: { bg: '#0A0A0B', panel: '#141416', border: 'rgba(212,175,55,.18)', soft: 'rgba(212,175,55,.10)', text: '#F5F0E4', dim: 'rgba(245,240,228,.82)', faint: 'rgba(245,240,228,.60)', accent: '#D4AF37', accent2: '#E8C96A',
      pri: 'linear-gradient(120deg,#F1D77C 0%,#D4AF37 55%,#B8860B 100%)', priInk: '#1A1400', hero: 'linear-gradient(135deg,#141416,#33290F)' }
  }[bt];
  var dark = bt !== 'light';
  var F = 'Inter,system-ui,-apple-system,"Segoe UI",Arial,sans-serif';
  var R = 'html.dr-embed';
  var L = function (list) { return list.map(function (s) { return R + ' ' + s; }).join(','); };   // one rule for many selectors
  var css = document.createElement('style'); css.id = 'drEmbed';
  css.textContent = [
    R + '{--bg:' + T.bg + ';--surface:' + T.panel + ';--line:' + T.border + ';--line-2:' + T.soft + ';--ink:' + T.text + ';--ink-2:' + T.dim + ';--ink-3:' + T.faint +
      ';--navy:' + T.text + ';--navy-2:' + T.dim + ';--steel:' + T.accent + ';--steel-soft:' + T.soft + ';--brand1:' + T.accent + ';--brand2:' + T.accent2 + ';--brand3:' + T.accent2 +
      ';--grad:' + T.pri + ';--sh-1:none;--sh-2:none;--sh-3:0 10px 30px rgba(0,0,0,.14);}',
    // status colours that stay readable on the dark themes
    dark ? R + '{--great:#3DDC97;--great-bg:rgba(61,220,151,.16);--ok:#3DDC97;--ok-bg:rgba(61,220,151,.12);--warn:#EAB308;--warn-bg:rgba(234,179,8,.14);--miss:#F87171;--miss-bg:rgba(248,113,113,.14);--info:#9FB0C2;--info-bg:rgba(159,176,194,.12);}' : '',
    // one flat background and Brahmastra's type in place of the page's own gradients and Poppins
    R + ' .app,' + R + ' .auth-wrap{background:' + T.bg + ' !important;font-family:' + F + ';}',
    R + ' body,' + R + ' .app *,' + R + ' .auth-wrap *{font-family:' + F + ';}',
    R + ' body{background:' + T.bg + ';}',
    // its own branding: Brahmastra already says where you are
    R + ' .nav .brand,' + R + ' .auth .brand,' + R + ' .auth > svg,' + R + ' .auth .sub,' + R + ' .fb-tab{display:none !important;}',
    R + ' .nav{background:' + T.panel + ';border-bottom:1px solid ' + T.border + ';backdrop-filter:none;-webkit-backdrop-filter:none;}',
    // a thin scrollbar in the theme's colours, so the page edge reads as Brahmastra's own
    R + '{scrollbar-width:thin;scrollbar-color:' + T.border + ' transparent;}',
    R + ' ::-webkit-scrollbar{width:10px;height:10px;}' + R + ' ::-webkit-scrollbar-thumb{background:' + T.border + ';border-radius:8px;border:2px solid ' + T.bg + ';}' + R + ' ::-webkit-scrollbar-track{background:transparent;}',
    // once Brahmastra draws the tabs itself (see below), Drishti's whole bar goes: tabs, name, role and Sign out.
    // Brahmastra signs people in and out of Drishti with its own session.
    'html.dr-lift .nav{display:none !important;}',
    // the page lines up with Brahmastra's own left edge instead of centring itself in a narrower column
    R + ' .wrap{max-width:none;padding:18px 2px 60px;animation:none;}',
    // Brahmastra's flat cards: one panel colour and its corners in place of the dot pattern, the white gradients and the lift
    L(['.panel', '.card', '.scopework-card', '.report-card', '.scope-block', '.volcard', '.tile', '.ov-card', '.ov-card.ok', '.ov-card.warn', '.ov-card.miss', '.ov-card.info',
      '.strip', '.auth', '.modal', '.bulkzone', '.od-bar', '.mapper', '.ov-dd']) + '{background:' + T.panel + ';box-shadow:none;border-radius:10px;animation:none;}',
    L(['.od-bar', '.mapper', '.ov-dd']) + '{border:1px solid ' + T.border + ';}',
    L(['.ov-card:hover', '.tile:hover', '.bulkzone.on', '.scopepick label:hover']) + '{transform:none;box-shadow:none;}',
    // the shimmer sweeps, spinning rings and running lines go: Brahmastra's pages hold still
    L(['.ov-card::after', '.scopework-card::after', '.report-card::after', '.scope-block::after', '.volcard::after', '.tile::after', '.btn.primary::after', '.runline::after',
      '.nav::after', '.report-card::before', '.bulkzone::before', '.fb-panel::before', '.shout::after', '.ov-topline']) + '{display:none;}',
    R + ' .section-h h2{background:transparent;box-shadow:none;padding:0;}',
    // the white boxes inside cards take the theme's soft and panel colours
    L(['.scope-head', '.scope-head:hover', '.hist th', '.hist th.hm', '.hist-g td', '.hist-g td.hm', '.hist-chart-row td', '.hist td.gap', 'table.hist th.gap', '.ownertag']) + '{background:' + T.soft + ';}',
    L(['.ov-dd-btn', '.ov-streak', '.br-sel', '.fm-search', '.fm-chip', '.fm-how-grid>div', '.fm-table', '.fm-sla-item', '.od-file', '.gd-steps>li', '.gd-table', '.gd-flow span', '.bfp-item']) +
      '{background:' + T.panel + ';border-color:' + T.border + ';box-shadow:none;}',
    // Brahmastra's buttons and fields: its corners and its primary colours
    R + ' .btn{border-radius:8px;box-shadow:none;}' + R + ' .btn:hover{transform:none;box-shadow:none;}',
    R + ' .btn.primary{color:' + T.priInk + ';box-shadow:none;}',
    L(['select', 'input', 'textarea']) + '{border-radius:8px;}',
    L(['.pp-month.on', '.pp-chip.on', '.fm-chip.on', '.calday.edge', '.pp-week.on .calday']) + '{color:' + T.priInk + ';}',
    R + ' .ov-hero{background:' + T.hero + ';border-radius:10px;}',
    // fixed dark lettering that would vanish on the dark themes
    dark ? L(['.note-box', '.hist-c.warn', '.ov-mval.warn']) + '{color:var(--warn);}' + L(['.hist-c.ok']) + '{color:var(--ok);}' + L(['.hist-c.miss', '.ov-streak']) + '{color:var(--miss);}' : ''
  ].filter(Boolean).join('\n');
  (document.head || root).appendChild(css);

  /* Brahmastra can draw Drishti's tabs (Overview, Clients, Formulas...) in its own tab row. It says so with a
     'brahmastra:hello' message; from then on the page sends its tabs ({id, label, on}) whenever they change, hides
     its own bar, and opens a tab when Brahmastra sends 'brahmastra:go'. An older Brahmastra never says hello, so
     nothing is hidden and Drishti's own bar stays. */
  var TAB_SEL = '.nav .tab';
  var lifted = false, last = '';
  var tabsNow = function () {
    return Array.prototype.map.call(document.querySelectorAll(TAB_SEL), function (b) {
      var l = b.textContent.trim(); return { id: l, label: l, on: b.classList.contains('on') };
    }).filter(function (t) { return t.label; });
  };
  var report = function () {
    if (!lifted || window.parent === window) return;
    var t = tabsNow(), k = JSON.stringify(t);
    if (k === last) return;
    last = k;
    window.parent.postMessage({ type: 'embed:tabs', app: 'drishti', tabs: t }, '*');
  };
  window.addEventListener('message', function (e) {
    if (e.source !== window.parent) return;
    var d = e.data || {};
    if (d.type === 'brahmastra:hello') { lifted = true; root.classList.add('dr-lift'); last = ''; report(); }
    else if (d.type === 'brahmastra:go' && d.id != null) {
      var hit = Array.prototype.filter.call(document.querySelectorAll(TAB_SEL), function (b) { return b.textContent.trim() === String(d.id); })[0];
      if (hit) { hit.click(); window.scrollTo(0, 0); }
    }
  });
  var watch = function () { new MutationObserver(report).observe(document.body, { subtree: true, childList: true, attributes: true, attributeFilter: ['class'] }); };
  if (document.body) watch(); else document.addEventListener('DOMContentLoaded', watch);
  // Brahmastra drops Drishti's sign-in token into this browser (same web address). If it arrives while
  // the welcome card is up, come in without asking anyone to click.
  window.addEventListener('storage', function (e) {
    if (e && e.key === 'drishti:token:v1' && e.newValue && document.querySelector('.auth-wrap')) location.reload();
  });
})();
