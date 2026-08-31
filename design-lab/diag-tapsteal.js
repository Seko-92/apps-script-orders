// =====================================================================================
// diag-tapsteal.js — DOES A POLL LANDING MID-TAP STEAL THE PICK?
//
// THE HYPOTHESIS (2026-08-19, from the picker's "maybe one time it didn't register"):
//   reconcileInto() calls fillPickRow() for EVERY row on EVERY tick, unconditionally,
//   and fillPickRow opens with `li.innerHTML = ''`. So every 20s each pick button is
//   destroyed and rebuilt even when nothing changed. A click only exists if pointer
//   down and up resolve to the same element — so a tick landing between the finger
//   going down and coming up can erase the button mid-gesture and no click is ever
//   generated. The delegated handler is fine; there is simply nothing for it to hear.
//
//   tap window ~100ms / poll 20000ms ≈ 0.5% per tap ≈ ~1 lost tap a day at this volume,
//   which is exactly what the floor reported.
//
// ⚠ THIS USES REAL MOUSE EVENTS AT REAL COORDINATES, never synthetic dispatch. If the
//   harness fabricated the click it would be proving its own arithmetic, not the board's
//   behaviour — the browser has to be the one that decides whether a click happened.
//
// Run against an older board to prove the fix:
//   git show HEAD:FloorBoard.html > /tmp/board-head.html
//   BOARD_FILE=/tmp/board-head.html node diag-tapsteal.js
// =====================================================================================
'use strict';
const fs=require('fs'), path=require('path');
const { chromium } = require('playwright');
const BOARD = process.env.BOARD_FILE || path.join(__dirname,'..','FloorBoard.html');
const TAG   = process.env.TAG || (process.env.BOARD_FILE ? 'HEAD' : 'working tree');
const MOCK  = require('./mock-tick.js');

let pass=0, fail=0;
const ok=(n,c,x)=>{ c?(pass++,console.log('  ✓ '+n)):(fail++,console.log('  ✗ '+n+(x!==undefined?'  → '+JSON.stringify(x):''))); };

(async()=>{
  const html=fs.readFileSync(BOARD,'utf8');
  const browser=await chromium.launch();
  const ctx=await browser.newContext({viewport:{width:1280,height:800},hasTouch:true,timezoneId:'America/Chicago'});
  const page=await ctx.newPage();

  const writes=[];                       // every boardStatus that reaches the server
  await page.route('http://hqlab.test/**', route=>{
    const url=route.request().url();
    if(url.includes('/api/board')){
      const body=JSON.parse(route.request().postData()||'{}');
      if(body.action==='boardStatus') writes.push(body);
      let res={ok:false,message:'unknown '+body.action};
      if(body.action==='boardTick')   res=Object.assign({ok:true},MOCK);
      if(body.action==='boardStatus') res={ok:true};
      if(body.action==='boardRadio')  res={ok:true,nowPlaying:''};
      return route.fulfill({contentType:'application/json',body:JSON.stringify(res)});
    }
    return route.fulfill({contentType:'text/html; charset=utf-8',body:html});
  });
  await page.route(/aladhan|open-meteo/, r=>r.abort());
  await page.goto('http://hqlab.test/',{waitUntil:'load'});
  await page.waitForFunction(()=>!document.getElementById('board').classList.contains('booting'),null,{timeout:20000}).catch(()=>{});
  await page.waitForTimeout(2400);

  // ⚠ DO NOT "HELP" THE BOARD HERE. The mock tick already carries
  // picker:'Shipping - Hatem 7788', so the gate passes on its own. My first cut
  // set pickerOverride={value:…} — but the board reads .name, so it fed the gate
  // `undefined` and every case failed in a way that looked exactly like a board
  // bug. Same trap diag-offline records for context.setOffline(). If the control
  // in case A fails, suspect the harness before the board.
  const pickerSeen = await page.evaluate(()=>currentPickerName());
  if(!pickerSeen){ console.log('  ⚠ no picker on the tick — every case would fail on the gate'); }

  /** Begin a tap on the first ✓ Pick, optionally repaint mid-gesture, finish the tap. */
  async function tap(repaintMidGesture){
    writes.length=0;
    const box=await page.evaluate(()=>{
      const b=document.querySelector('.pick-do');
      if(!b) return null;
      const r=b.getBoundingClientRect();
      return {x:r.x+r.width/2, y:r.y+r.height/2, order:b.getAttribute('data-order'), sku:b.getAttribute('data-sku')};
    });
    if(!box) return {err:'no ✓ Pick button on screen'};
    await page.mouse.move(box.x, box.y);
    await page.mouse.down();                         // finger down
    if(repaintMidGesture){
      // exactly what a poll does — repaint from the tick already in hand
      await page.evaluate(()=>{
        paintPickList(applyPickOverrides((lastTick.openOrders||[]).slice()),
                      lastTick.openOrdersTotal, lastTick.openOrdersBy);
      });
    }
    await page.mouse.up();                           // finger up
    await page.waitForTimeout(450);
    return {target:box, wrote:writes.length, sent:writes[0]||null};
  }

  console.log('\n══ diag-tapsteal · '+TAG+' ══');

  console.log('\nA · CONTROL — a clean tap, nothing interrupting it');
  let r=await tap(false);
  ok('the pick reaches the server exactly once', r.wrote===1, r);
  ok('and it names the row the picker actually touched',
     r.sent && r.sent.orderId===r.target.order && r.sent.sku===r.target.sku, r.sent);

  // put the board back to a clean state between cases
  await page.evaluate(()=>{ try{ pickOverrides={}; }catch(e){} });
  await page.waitForTimeout(400);

  console.log('\nB · THE REPORTED CASE — a poll repaints while the finger is down');
  r=await tap(true);
  ok('the pick STILL reaches the server (the tap is not swallowed)', r.wrote===1, r);
  ok('and it is still the right row', r.wrote===1 && r.sent && r.sent.sku===r.target.sku, r.sent);

  await page.evaluate(()=>{ try{ pickOverrides={}; }catch(e){} });
  await page.waitForTimeout(400);

  console.log('\nC · REGRESSION NET — a repaint AFTER the tap must not double-fire');
  writes.length=0;
  const box=await page.evaluate(()=>{const b=document.querySelector('.pick-do');const q=b.getBoundingClientRect();return{x:q.x+q.width/2,y:q.y+q.height/2};});
  await page.mouse.click(box.x, box.y);
  await page.waitForTimeout(300);
  await page.evaluate(()=>{ paintPickList(applyPickOverrides((lastTick.openOrders||[]).slice()), lastTick.openOrdersTotal, lastTick.openOrdersBy); });
  await page.waitForTimeout(400);
  ok('exactly one write, never two', writes.length===1, writes.length);

  console.log('\nD · THE HOLD — a repaint asked for mid-gesture must be DEFERRED, not dropped');
  await page.evaluate(()=>{ try{ pickOverrides={}; }catch(e){} });
  await page.waitForTimeout(300);
  const held = await page.evaluate(async ()=>{
    const b=document.querySelector('.pick-do'); if(!b) return {err:'no button'};
    b.setAttribute('data-probe','1');
    // simulate a finger going down on the list, then a tick arriving
    document.getElementById('pickSplit').dispatchEvent(new PointerEvent('pointerdown',{bubbles:true}));
    paintPickList(applyPickOverrides((lastTick.openOrders||[]).slice()),
                  lastTick.openOrdersTotal, lastTick.openOrdersBy);
    const duringHold = !!(document.querySelector('.pick-do')||{}).getAttribute
                        && !!document.querySelector('.pick-do[data-probe="1"]');
    window.dispatchEvent(new PointerEvent('pointerup',{bubbles:true}));
    await new Promise(r=>setTimeout(r,120));
    const afterRelease = !!document.querySelector('.pick-do[data-probe="1"]');
    return { survivedDuringHold: duringHold, rebuiltAfterRelease: !afterRelease };
  });
  ok('the button is untouched while the finger is down', held.survivedDuringHold===true, held);
  ok('and the held repaint runs as soon as the finger lifts', held.rebuiltAfterRelease===true, held);

  console.log('\nE · THE MECHANISM (informational) — the churn that made this possible');
  const churn=await page.evaluate(()=>{
    const before=document.querySelector('.pick-do');
    if(!before) return {err:'no button'};
    before.setAttribute('data-probe','1');
    paintPickList(applyPickOverrides((lastTick.openOrders||[]).slice()),
                  lastTick.openOrdersTotal, lastTick.openOrdersBy);
    const after=document.querySelector('.pick-do');
    return { sameNode: before===after,
             markSurvived: !!(after && after.getAttribute('data-probe')),
             stillAttached: document.contains(before) };
  });
  console.log('     button node survives a repaint: '+churn.sameNode+
              '   (mark survived: '+churn.markSurvived+', old node still in the DOM: '+churn.stillAttached+')');
  console.log('     ⚠ every tick still rebuilds every row even when nothing changed —');
  console.log('       that is the churn the hold now protects the tap from. Reducing it');
  console.log('       (skip the rebuild when a row is unchanged) is a separate, safe win.');

  console.log('\n'+(fail===0?'✅ ':'❌ ')+pass+' passed, '+fail+' failed   ['+TAG+']');
  await browser.close();
  process.exit(fail===0?0:1);
})().catch(e=>{ console.error('CRASH: '+e.message); process.exit(1); });
