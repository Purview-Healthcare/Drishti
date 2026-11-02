/* Inside Brahmastra: ?embed=1&theme=light|navy|gold
   My day & reports and the SLA tab open Drishti with these parameters. The page then hides its own
   branding (the nav logo, the welcome card's logo and byline, the Feedback tab) and takes Brahmastra's
   colours and type, so it reads as one app. Opened directly, without ?embed=1, nothing here runs and
   Drishti looks exactly as before. Numbers, data and controls are untouched either way. */
(function () {
  'use strict';
  var q; try { q = new URLSearchParams(location.search); } catch (e) { return; }
  if (q.get('embed') !== '1') return;
  var bt = { light: 1, navy: 1, gold: 1 }[q.get('theme')] ? q.get('theme') : 'light';
  var root = document.documentElement;
  root.classList.add('dr-embed'); root.setAttribute('data-bt', bt);
  var T = {
    light: { bg: '#EEF1F4', panel: '#FFFFFF', panel2: '#F5F7F9', border: '#DFE4EA', soft: '#EDF1F4', text: '#0F1B24', dim: '#3B4A55', faint: '#627079', accent: '#0E8C74', accent2: '#12B886' },
    navy: { bg: '#071E33', panel: '#0B2740', panel2: '#09223A', border: 'rgba(255,255,255,.12)', soft: 'rgba(255,255,255,.07)', text: '#FFFFFF', dim: 'rgba(255,255,255,.82)', faint: 'rgba(255,255,255,.60)', accent: '#2BD4A4', accent2: '#38C8E8' },
    gold: { bg: '#0A0A0B', panel: '#141416', panel2: '#101012', border: 'rgba(212,175,55,.18)', soft: 'rgba(212,175,55,.10)', text: '#F5F0E4', dim: 'rgba(245,240,228,.82)', faint: 'rgba(245,240,228,.60)', accent: '#D4AF37', accent2: '#E8C96A' }
  }[bt];
  var dark = bt !== 'light';
  var F = 'Inter,system-ui,-apple-system,"Segoe UI",Arial,sans-serif';
  var R = 'html.dr-embed';
  var css = document.createElement('style'); css.id = 'drEmbed';
  css.textContent = [
    R + '{--bg:' + T.bg + ';--surface:' + T.panel + ';--line:' + T.border + ';--line-2:' + T.soft + ';--ink:' + T.text + ';--ink-2:' + T.dim + ';--ink-3:' + T.faint +
      ';--navy:' + T.text + ';--navy-2:' + T.dim + ';--steel:' + T.accent + ';--steel-soft:' + T.soft +
      ';--grad:linear-gradient(120deg,' + T.accent + ',' + T.accent2 + ');--sh-1:none;--sh-2:none;--sh-3:0 10px 30px rgba(0,0,0,.14);}',
    // status colours that stay readable on the dark themes
    dark ? R + '{--great:#3DDc97;--great-bg:rgba(61,220,151,.16);--ok:#3DDC97;--ok-bg:rgba(61,220,151,.12);--warn:#EAB308;--warn-bg:rgba(234,179,8,.14);--miss:#F87171;--miss-bg:rgba(248,113,113,.14);--info:#9FB0C2;--info-bg:rgba(159,176,194,.12);}' : '',
    // one flat background and Brahmastra's type in place of the page's own gradients and Poppins
    R + ' .app,' + R + ' .auth-wrap{background:' + T.bg + ' !important;font-family:' + F + ';}',
    R + ' body,' + R + ' .app *,' + R + ' .auth-wrap *{font-family:' + F + ';}',
    R + ' body{background:' + T.bg + ';}',
    // its own branding: Brahmastra already says where you are
    R + ' .nav .brand,' + R + ' .auth .brand,' + R + ' .auth > svg,' + R + ' .auth .sub,' + R + ' .fb-tab{display:none !important;}',
    R + ' .nav{background:' + T.panel + ';border-bottom:1px solid ' + T.border + ';backdrop-filter:none;-webkit-backdrop-filter:none;}',
    R + ' .auth{box-shadow:0 10px 30px rgba(0,0,0,.12);}',
    // a thin scrollbar in the theme's colours, so the page edge reads as Brahmastra's own
    R + '{scrollbar-width:thin;scrollbar-color:' + T.border + ' transparent;}',
    R + ' ::-webkit-scrollbar{width:10px;height:10px;}' + R + ' ::-webkit-scrollbar-thumb{background:' + T.border + ';border-radius:8px;border:2px solid ' + T.bg + ';}' + R + ' ::-webkit-scrollbar-track{background:transparent;}',
    // once Brahmastra draws the tabs itself (see below), Drishti's whole bar goes: tabs, name, role and Sign out.
    // Brahmastra signs people in and out of Drishti with its own session.
    'html.dr-lift .nav{display:none !important;}',
    // the page lines up with Brahmastra's own left edge instead of centring itself in a narrower column
    R + ' .wrap{max-width:none;padding:18px 2px 60px;}'
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
