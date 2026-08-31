const { chromium } = require('playwright');
const path=require('path');
const fs=require('fs');
// ⚠ SELF-SUFFICIENT: builds its own resolved copy from the REAL Sidebar.html
// instead of a pre-made file in a session scratchpad that no longer exists.
// Served over http:// because setContent/file:// give an origin where
// localStorage THROWS, and the sidebar reads it at the top of its script —
// one throw there kills every statement after it.
// ⚠ SIDEBAR_SRC lets this run against an older revision to PROVE the assertions
//   bite:  git show HEAD:Sidebar.html > /tmp/old.html && SIDEBAR_SRC=/tmp/old.html node check-sidebar.js
const SRC=process.env.SIDEBAR_SRC||path.join(__dirname,'..','Sidebar.html');
const HTML=fs.readFileSync(SRC,'utf8').replace("'<?!= boardApiUrl ?>'","''");

(async()=>{
  const b=await chromium.launch();
  const p=await b.newPage({viewport:{width:310,height:2600},deviceScaleFactor:2});
  const errs=[]; p.on('pageerror',e=>errs.push(String(e)));
  // Stub the server. Without getDisplayUrls the Displays card renders empty
  // and three assertions fail for a reason that has nothing to do with the page.
  await p.addInitScript(()=>{
    const D={getDisplayUrls:{board:'https://hq.yassinqurabi.com/',wall:'https://hq.yassinqurabi.com/wall',hosted:true},
             getSidebarTick:null,getCurrentPicker:'',getActionableAlerts:null};
    const mk=(su,fa)=>new Proxy({},{get(_,k){
      if(k==='withSuccessHandler')return f=>mk(f,fa);
      if(k==='withFailureHandler')return f=>mk(su,f);
      return ()=>{const v=Object.prototype.hasOwnProperty.call(D,k)?D[k]:null; if(su)setTimeout(()=>su(v),0);};}});
    window.google={script:{run:mk(null,null),host:{close(){},setHeight(){}},url:{getLocation(f){f({parameter:{}});}}}};
  });
  p.on('console',m=>{if(m.type()==='error')errs.push(m.text());});
  await p.route('http://hq.test/**',r=>r.fulfill({contentType:'text/html; charset=utf-8',body:HTML}));
  await p.goto('http://hq.test/sidebar');
  await p.waitForTimeout(2400);
  let fail=0; const ok=(n,c,x)=>{c?console.log('  ok  '+n):(fail++,console.log('  XX  '+n+(x!==undefined?'  -> '+JSON.stringify(x):'')));};
  await p.evaluate(()=>document.querySelectorAll('#modules details').forEach(d=>d.open=true));
  /* ⚠ THE PALETTE HAS NO ROWS UNTIL IT IS OPENED — #cmdList is populated by
     renderCommands() on open, so a check that does not do this reads an empty
     list and passes vacuously. That blindness is why the palette outlived two
     emoji sweeps. */
  await p.evaluate(()=>{ if (typeof openPalette==='function') openPalette(); });
  await p.waitForTimeout(250);
  const r=await p.evaluate(()=>{
    const cs=n=>getComputedStyle(n);
    const card=document.querySelector('.card[data-id="displays"]');
    const zoneOf=id=>{const c=document.querySelector(`.card[data-id="${id}"]`);return c?c.dataset.zone:null;};
    return {cards:document.querySelectorAll('.card').length,
      displays:!!card, displaysZone:zoneOf('displays'),
      /* the 08-28 door (2026-08-29) — ⚠ an unmapped data-id falls back to the
         generic m-board icon SILENTLY, so the card needs its own net, not just
         the aggregate counts. */
      mlCard:!!document.querySelector('.card[data-id="missing-line"]'),
      mlZone:zoneOf('missing-line'),
      mlMark:(document.querySelector('.card[data-id="missing-line"] .card-icon use')||{})
              .getAttribute && document.querySelector('.card[data-id="missing-line"] .card-icon use')
              .getAttribute('href'),
      mlBtnEmoji:[...document.querySelectorAll('.card[data-id="missing-line"] button')]
        .filter(b=>/[\u{1F300}-\u{1FAFF}\u{2600}-\u{27BF}]/u.test(b.innerText)).length,
      reachGone:!document.querySelector('[data-id="reach-probe"]'),
      fbGone:!document.querySelector('[data-id="floor-board"]'),
      boardHref:(document.getElementById('dsBoardLink')||{}).href,
      wallHref:(document.getElementById('dsWallLink')||{}).href,
      wallShown:cs(document.getElementById('dsWallLink')).display!=='none',
      urlBox:(document.getElementById('dsUrl')||{}).value,
      btnMarks:document.querySelectorAll('#modules button svg').length,
      btnBroken:[...document.querySelectorAll('#modules button svg use')]
        .filter(u=>!document.querySelector(u.getAttribute('href'))).map(u=>u.getAttribute('href')),
      btnEmoji:[...document.querySelectorAll('#modules button')]
        .filter(b=>/[\u{1F300}-\u{1FAFF}\u{2600}-\u{27BF}]/u.test(b.innerText))
        .map(b=>b.innerText.trim().replace(/\s+/g,' ').slice(0,30)),
      emoji:[...document.querySelectorAll('.card-icon,.ctrl-btn')]
        .filter(e=>/[\u{1F300}-\u{1FAFF}\u{2600}-\u{27BF}]/u.test(e.textContent)).length,
      marks:document.querySelectorAll('.card-icon svg,.ctrl-btn svg').length,
      upper:[...document.querySelectorAll('.card-title h3')].filter(t=>cs(t).textTransform==='uppercase').length,
      snakeInPalette:document.body.innerHTML.indexOf('Launch Snake Game')>=0,
      /* ── the command palette + the picker banner (2026-08-20). Both were blind
         spots for opposite reasons: the palette renders into #cmdList only once
         OPENED and its icons come from a JS array, and the picker banner sits
         OUTSIDE #modules so the _hqMarks pass never walked it. Neither was
         covered by any net above, which is exactly how both outlived two emoji
         sweeps. openPalette() is called by the reader below before this runs. */
      cmdRows:document.querySelectorAll('#cmdList .cmd').length,
      cmdMarks:document.querySelectorAll('#cmdList .cmd-icon svg').length,
      cmdBroken:[...document.querySelectorAll('#cmdList .cmd-icon use')]
        .filter(u=>!document.querySelector(u.getAttribute('href'))).map(u=>u.getAttribute('href')),
      cmdEmoji:[...document.querySelectorAll('#cmdList .cmd')]
        .filter(e=>/[\u{1F300}-\u{1FAFF}\u{2600}-\u{27BF}]/u.test(e.textContent)).length,
      pickerMark:!!document.querySelector('#pickerBanner .picker-icon svg use'),
      pickerEmoji:/[\u{1F300}-\u{1FAFF}\u{2600}-\u{27BF}]/u
        .test((document.getElementById('pickerBanner')||{textContent:''}).textContent),
      zones:[...document.querySelectorAll('.zone-header,.zone-title')].map(e=>e.textContent.trim()),
      // ── the lock sequence must be reachable WITHOUT the Apps Script editor (2026-08-30).
      //    The n8n account was the one step that lived in code; it is a field now.
      lockAcctInput:!!document.getElementById('lockAcct'),
      lockAcctSave:[...document.querySelectorAll('.card[data-id="sheet-protect"] button')]
        .some(b=>/saveN8nAccountNow/.test(b.getAttribute('onclick')||'')),
      lockAcctResult:!!document.getElementById('lockAcctResult'),
      lockHeaderLoads:/lockPanelLoad/.test(
        (document.querySelector('.card[data-id="sheet-protect"] .card-header')||{})
          .getAttribute?.('onclick')||''),
      lockBtns:[...document.querySelectorAll('.card[data-id="sheet-protect"] button')]
        .map(b=>b.getAttribute('onclick')||'')
        .filter(o=>/verifyOwnerBridge|lockAllOrdersNow|describeAllOrdersLock|unprotectAllOrdersSheet/.test(o))
        .length};
  });
  console.log('  ',JSON.stringify(r,null,0).slice(0,400));
  ok('27 cards', r.cards===27, r.cards);
  ok('Displays card exists', r.displays);
  ok('Displays is in TODAY\'S WORK', r.displaysZone==='today', r.displaysZone);
  ok('missing-line card exists', r.mlCard);
  ok('missing-line is in TODAY\'S WORK', r.mlZone==='today', r.mlZone);
  ok('missing-line wears m-case, not the m-board fallback', r.mlMark==='#m-case', r.mlMark);
  ok('missing-line buttons carry NO emoji', r.mlBtnEmoji===0, r.mlBtnEmoji);
  ok('floor-board card gone', r.fbGone);
  ok('reach-probe card gone', r.reachGone);
  ok('board link points at the HOSTED board', /hq\.yassinqurabi\.com\/$/.test(r.boardHref||''), r.boardHref);
  ok('wall link points at /wall', /hq\.yassinqurabi\.com\/wall$/.test(r.wallHref||''), r.wallHref);
  ok('wall link is shown when hosted', r.wallShown);
  ok('copy box holds the board url', (r.urlBox||'').indexOf('hq.yassinqurabi.com')>0, r.urlBox);
  ok('zero emoji icons', r.emoji===0, r.emoji);
  // ── button marks (2026-08-20). v3.6 did the card HEADS and stopped; the
  // buttons inside the cards still carried 39 emoji. ⚠ Three markup shapes and
  // an astral-surrogate trap made this a 3-pass fix, so it gets a real net.
  ok('button marks rendered',r.btnMarks>=35,r.btnMarks);
  ok('no unmapped <use> on a button',r.btnBroken.length===0,r.btnBroken);
  ok('only the arcade keeps an emoji',r.btnEmoji.length===1&&/Arcade/.test(r.btnEmoji[0]),r.btnEmoji);
  ok('marks rendered', r.marks>=29, r.marks);
  ok('all card titles uppercase', r.upper===27, r.upper);
  ok('Snake gone from the palette', !r.snakeInPalette);
  // ── the last two emoji surfaces inside the sidebar
  ok('palette renders its commands', r.cmdRows>=17, r.cmdRows);
  ok('every palette row wears a drawn mark', r.cmdMarks===r.cmdRows, [r.cmdMarks,r.cmdRows]);
  ok('no dangling <use> in the palette', r.cmdBroken.length===0, r.cmdBroken);
  ok('ZERO emoji left in the palette', r.cmdEmoji===0, r.cmdEmoji);
  ok('the picker banner wears a mark, not 👤', r.pickerMark && !r.pickerEmoji,
     {mark:r.pickerMark, emoji:r.pickerEmoji});
  ok('four zones intact', r.zones.length===4, r.zones);
  // ── ⭐ THE WHOLE LOCK SEQUENCE LIVES IN THE PANEL (2026-08-30). The n8n account used to
  //    mean opening the editor and editing `var VALUE` in code; installAllOrdersLock()
  //    refuses without it, so that one value forced the entire sequence out of the sidebar.
  ok('the n8n account is a FIELD, not an editor edit', r.lockAcctInput);
  ok('...wired to the owner-gated setter', r.lockAcctSave);
  ok('...with somewhere to report the result', r.lockAcctResult);
  ok('the card reads its state when opened', r.lockHeaderLoads);
  ok('all four lock controls present', r.lockBtns===4, r.lockBtns);
  /* ⚠⚠ NO CARD MAY BE CLIPPED BY ITS OWN max-height.
     .card-body caps at 800px with overflow:hidden, so a card that outgrows it loses its
     bottom SILENTLY — the elements stay in the DOM, so every markup assertion above still
     passes. Sheet Protection crossed it the moment the lock controls shipped (2026-08-30)
     and hid its identity buttons; only a render showed it. This is the net so the next one
     is loud. Cards are force-expanded and transitions killed, or clientHeight is measured
     mid-animation and the check reports nonsense. */
  const clipped = await p.evaluate(() => {
    document.querySelectorAll('.card.collapsed').forEach(c => c.classList.remove('collapsed'));
    document.querySelectorAll('.card-body').forEach(b => { b.style.transition = 'none'; });
    void document.body.offsetHeight;
    return [...document.querySelectorAll('.card')].map(c => {
      const bd = c.querySelector('.card-body');
      if (!bd) return null;
      return bd.scrollHeight > bd.clientHeight + 2
        ? (c.dataset.id || '?') + ' ' + bd.scrollHeight + 'px > ' + bd.clientHeight + 'px' : null;
    }).filter(Boolean);
  });
  ok('no card body is clipped by its max-height', clipped.length===0, clipped);

  /* ⚠⚠ EVERY var(--x) MUST RESOLVE.
     An UNDEFINED custom property is not a no-op and does not fall back to the previous
     rule — the declaration becomes "invalid at computed-value time" and resets to INITIAL.
     `--brand-yellow` was never defined, so .ml-seg button.on computed background:transparent
     and border-color:currentColor, and the selected Missing/Replacement segment rendered as
     a black outline instead of brand yellow. `--text-dim` did the same to two more rules.
     Both looked like styling choices, which is why they survived. */
  const badVars = await p.evaluate(() => {
    const defined = new Set(), used = new Set();
    for (const sh of document.styleSheets) {
      let rs; try { rs = sh.cssRules; } catch (e) { continue; }
      /* ⚠⚠ PROCESS THE RULE, THEN RECURSE — never `if (r.cssRules) { …; continue; }`.
         Modern Chrome gives EVERY CSSStyleRule a `cssRules` property (an empty list, for
         CSS Nesting), and an empty CSSRuleList is TRUTHY. That early-continue skipped all
         619 real rules and processed only @keyframes steps, so the check saw no definitions
         at all and reported a correctly-defined variable as missing. */
      const walk = list => { for (const r of list) {
        if (r.style) {
        // ⚠ Read DEFINITIONS from cssText, not r.style[i]. The indexed enumeration of a
        //   CSSStyleDeclaration does not reliably include custom properties, which made an
        //   earlier cut of this check report --shadow-md (defined twice, in the light and
        //   dark palettes) as missing. A net that cries wolf gets switched off.
          const txt = r.style.cssText || '';
          (txt.match(/(^|;)\s*(--[\w-]+)\s*:/g) || [])
            .forEach(x => defined.add(x.replace(/^[;\s]*/, '').replace(/\s*:$/, '')));
          (txt.match(/var\(\s*(--[\w-]+)/g) || [])
            .forEach(x => used.add(x.replace(/var\(\s*/, '')));
        }
        if (r.cssRules && r.cssRules.length) walk(r.cssRules);
      }};
      walk(rs);
    }
    return [...used].filter(v => !defined.has(v));
  });
  ok('every var(--x) resolves to a defined custom property', badVars.length===0, badVars);

  /* ⚠ NO FLEX ROW OF BUTTONS MAY BE MISALIGNED.
     The generic `button` rule sets margin-top:5px and only `button:first-child` escapes it,
     and .btn-secondary re-adds it with !important. In a flex row that pushes one button down
     and — under the default align-items:stretch — leaves it SHORTER than its neighbour.
     Measured 32px vs 27px on the Missing/Replacement segments and 30 vs 35 on its buttons. */
  const misaligned = await p.evaluate(() => {
    const out = [];
    document.querySelectorAll('#modules *').forEach(c => {
      const cs = getComputedStyle(c);
      if (!/flex/.test(cs.display) || cs.flexDirection.indexOf('column') === 0) return;
      const kids = [...c.children].filter(k => k.tagName === 'BUTTON');
      if (kids.length < 2) return;
      const box = kids.map(k => k.getBoundingClientRect());
      const tops = box.map(r => Math.round(r.top)), hs = box.map(r => Math.round(r.height));
      if (new Set(tops).size > 1 || new Set(hs).size > 1) {
        const card = c.closest('.card');
        out.push((card ? card.dataset.id : '?') + ' .' + String(c.className).split(' ')[0]
                 + ' tops=' + JSON.stringify(tops) + ' h=' + JSON.stringify(hs));
      }
    });
    return out;
  });
  ok('no flex row of buttons is misaligned', misaligned.length===0, misaligned);
  ok('no console errors', errs.length===0, errs.slice(0,3));
  await p.screenshot({path:path.join(__dirname,'renders','sidebar-final.png'),clip:{x:0,y:0,width:310,height:820}});
  console.log('\n'+(fail?('XX '+fail+' FAILED'):'OK all checks passed'));
  await b.close(); process.exit(fail?1:0);
})();
