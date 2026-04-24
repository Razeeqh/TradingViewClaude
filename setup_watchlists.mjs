/**
 * Adds curated NSE high-conviction scalp candidates to the existing (cleared) watchlist.
 * Stocks selected on 3 criteria: VCP/breakout pattern + positive news catalyst + F&O liquidity.
 * Market context symbols placed first.
 */
import { connect, evaluate, getClient } from './src/connection.js';

await connect();
const c = await getClient();
const sleep = ms => new Promise(r => setTimeout(r, ms));

const WATCHLIST = [
  // ── Market Context (always watch these first) ─────────────────────────────
  'NSE:NIFTY',
  'NSE:BANKNIFTY',
  // ── Tier 1: VCP Breakout + Strong Catalyst (72-70/100) ───────────────────
  'NSE:COALINDIA',   // VCP breakout +3.61% | Q4 results + dividend Apr 27
  'NSE:NESTLEIND',   // 52-wk high breakout ₹1431 | Q4 profit +26% YoY
  // ── Tier 2: Post-Earnings Base + Institutional (66-64/100) ───────────────
  'NSE:ICICIBANK',   // Q4 beat 5.8% | record-low GNPA | ₹12 dividend
  'NSE:HDFCBANK',    // Q4 beat | loans +12% | ₹13 dividend
  'NSE:BEL',         // Defence | FY26 rev +16% | order book ₹74,000 cr
  // ── Tier 3: Pre-Results VCP + Sales Momentum (62-60/100) ─────────────────
  'NSE:BAJFINANCE',  // AUM ₹5L cr milestone | new loans +20.5% | results Apr 29
  'NSE:BAJAJ-AUTO',  // Sales +20% | EPS beat | results May 6
  'NSE:AXISBANK',    // Q4 results out | banking sector momentum
  // ── Tier 4: Sector Momentum + PSU Catalyst (58-55/100) ───────────────────
  'NSE:SBIN',        // PSU banking tailwind
  'NSE:JSWSTEEL',    // Metals sector momentum | infrastructure spend
  'NSE:POWERGRID',   // PSU energy | renewable capex | high dividend
  'NSE:NTPC',        // Renewable expansion | energy security theme
];

// ── Helpers ──────────────────────────────────────────────────────────────────

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
  const addClicked = await evaluate(`
    (function() {
      var selectors = [
        '[data-name="add-symbol-button"]',
        '[aria-label="Add symbol"]',
        '[aria-label*="Add symbol"]',
        'button[class*="addSymbol"]',
      ];
      for (var i = 0; i < selectors.length; i++) {
        var el = document.querySelector(selectors[i]);
        if (el && el.offsetParent !== null) { el.click(); return { found: true }; }
      }
      var panel = document.querySelector('[class*="layout__area--right"]');
      if (panel) {
        var buttons = panel.querySelectorAll('button');
        for (var j = 0; j < buttons.length; j++) {
          if (buttons[j].textContent.trim() === '+' && buttons[j].offsetParent) {
            buttons[j].click(); return { found: true, method: 'plus' };
          }
        }
      }
      return { found: false };
    })()`);

  if (!addClicked?.found) {
    console.log(`  ⚠️  Add button not found — skipping ${symbol}`);
    return;
  }
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

// ── Main ─────────────────────────────────────────────────────────────────────

console.log('Opening watchlist panel...');
await ensureWatchlistOpen();

console.log(`\nAdding ${WATCHLIST.length} high-conviction NSE scalp candidates:\n`);
for (const sym of WATCHLIST) {
  await addSymbol(sym);
}

console.log(`\n🏁 Done — ${WATCHLIST.length} stocks added to watchlist.`);
console.log('\nWatchlist breakdown:');
console.log('  Market context : NIFTY, BANKNIFTY');
console.log('  Tier 1 (70-72) : COALINDIA, NESTLEIND');
console.log('  Tier 2 (64-66) : ICICIBANK, HDFCBANK, BEL');
console.log('  Tier 3 (60-62) : BAJFINANCE, BAJAJ-AUTO, AXISBANK');
console.log('  Tier 4 (55-58) : SBIN, JSWSTEEL, POWERGRID, NTPC');
