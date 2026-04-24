/**
 * Adds 5 new analyst-recommended stocks to the existing TradingView watchlist.
 * All have fresh buy calls from Motilal Oswal, ICICI Securities, or Emkay (April 2026).
 */
import { connect, evaluate, getClient } from './src/connection.js';

await connect();
const c = await getClient();
const sleep = ms => new Promise(r => setTimeout(r, ms));

const NEW_STOCKS = [
  'NSE:TRENT',     // Motilal Oswal Buy ₹5,250 — Apr 22
  'NSE:CIPLA',     // ICICI Securities Buy ₹1,550 — Apr 24
  'NSE:SBILIFE',   // Emkay + ICICI Sec Buy ₹2,345 — Apr 23
  'NSE:HDFCAMC',   // Motilal Oswal Buy ₹3,170 — Apr 17 (+5% post Q4)
  'NSE:CHOLAFIN',  // Motilal Oswal Buy +21% upside — Apr 16
];

async function ensureWatchlistOpen() {
  await evaluate(`
    (function() {
      var btn = document.querySelector('[data-name="base-watchlist-widget-button"]')
               || document.querySelector('[aria-label*="Watchlist"]');
      if (btn) {
        var active = btn.getAttribute('aria-pressed') === 'true'
                  || btn.className.indexOf('active') !== -1
                  || btn.className.indexOf('Active') !== -1;
        if (!active) btn.click();
      }
    })()`);
  await sleep(700);
}

async function addSymbol(symbol) {
  const clicked = await evaluate(`
    (function() {
      var selectors = ['[data-name="add-symbol-button"]','[aria-label="Add symbol"]','[aria-label*="Add symbol"]','button[class*="addSymbol"]'];
      for (var i = 0; i < selectors.length; i++) {
        var el = document.querySelector(selectors[i]);
        if (el && el.offsetParent !== null) { el.click(); return { found: true }; }
      }
      var panel = document.querySelector('[class*="layout__area--right"]');
      if (panel) {
        var btns = panel.querySelectorAll('button');
        for (var j = 0; j < btns.length; j++) {
          if (btns[j].textContent.trim() === '+' && btns[j].offsetParent) {
            btns[j].click(); return { found: true }; }
        }
      }
      return { found: false };
    })()`);
  if (!clicked?.found) { console.log(`  ⚠️  Skipped ${symbol}`); return; }
  await sleep(350);
  await c.Input.insertText({ text: symbol });
  await sleep(600);
  await c.Input.dispatchKeyEvent({ type: 'keyDown', key: 'Enter', code: 'Enter', windowsVirtualKeyCode: 13 });
  await c.Input.dispatchKeyEvent({ type: 'keyUp',   key: 'Enter', code: 'Enter', windowsVirtualKeyCode: 13 });
  await sleep(300);
  await c.Input.dispatchKeyEvent({ type: 'keyDown', key: 'Escape', code: 'Escape', windowsVirtualKeyCode: 27 });
  await c.Input.dispatchKeyEvent({ type: 'keyUp',   key: 'Escape', code: 'Escape', windowsVirtualKeyCode: 27 });
  await sleep(250);
  console.log(`  ✅  ${symbol}`);
}

await ensureWatchlistOpen();
console.log('\nAdding 5 analyst-backed stocks to TradingView watchlist:\n');
for (const sym of NEW_STOCKS) await addSymbol(sym);
console.log('\n🏁 Done — 5 stocks added.');
