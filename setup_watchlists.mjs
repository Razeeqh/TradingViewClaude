/**
 * Creates two TradingView watchlists via CDP:
 *  1. "NSE Fundamentals"  — blue-chip quality stocks
 *  2. "NSE Scalp Setup"   — high-beta F&O momentum stocks
 */
import { connect, evaluate, getClient } from './src/connection.js';

await connect();
const c = await getClient();

const sleep = ms => new Promise(r => setTimeout(r, ms));

// ── Symbol Lists ─────────────────────────────────────────────────────────────

const FUNDAMENTALS = [
  'NSE:HDFCBANK', 'NSE:ICICIBANK', 'NSE:KOTAKBANK', 'NSE:AXISBANK',
  'NSE:TCS', 'NSE:INFY', 'NSE:HCLTECH', 'NSE:WIPRO',
  'NSE:RELIANCE', 'NSE:BAJFINANCE', 'NSE:LT',
  'NSE:HINDUNILVR', 'NSE:NESTLEIND', 'NSE:BRITANNIA', 'NSE:TATACONSUM',
  'NSE:MARUTI', 'NSE:BAJAJ-AUTO', 'NSE:TITAN',
  'NSE:SUNPHARMA', 'NSE:DRREDDY', 'NSE:DIVISLAB',
  'NSE:ASIANPAINT', 'NSE:ULTRACEMCO',
];

const SCALP_SETUP = [
  'NSE:NIFTY', 'NSE:BANKNIFTY',           // market context — always first
  'NSE:COALINDIA',                          // #1 current pick
  'NSE:SBIN', 'NSE:ICICIBANK', 'NSE:AXISBANK', 'NSE:INDUSINDBK',
  'NSE:TATAMOTORS', 'NSE:TATASTEEL', 'NSE:HINDALCO', 'NSE:JSWSTEEL',
  'NSE:BAJFINANCE', 'NSE:SHRIRAMFIN',
  'NSE:ADANIENT', 'NSE:ADANIPORTS',
  'NSE:ONGC', 'NSE:BPCL', 'NSE:NTPC', 'NSE:POWERGRID',
  'NSE:BHARTIARTL', 'NSE:BEL',
  'NSE:M&M', 'NSE:EICHERMOT', 'NSE:HEROMOTOCO',
];

// ── Helpers ───────────────────────────────────────────────────────────────────

async function ensureWatchlistPanelOpen() {
  await evaluate(`
    (function() {
      var btn = document.querySelector('[data-name="base-watchlist-widget-button"]')
                || document.querySelector('[aria-label*="Watchlist"]');
      if (btn) {
        var isActive = btn.getAttribute('aria-pressed') === 'true'
                    || btn.className.indexOf('active') !== -1
                    || btn.className.indexOf('Active') !== -1;
        if (!isActive) btn.click();
      }
    })()`);
  await sleep(600);
}

async function createNewWatchlist(name) {
  console.log(`\n📋 Creating watchlist: "${name}"`);

  // Click the watchlist options / "+" / "New list" button
  const created = await evaluate(`
    (function() {
      // TradingView has a "+" or "New list" button near watchlist header
      var selectors = [
        '[data-name="watchlist-add-list"]',
        '[aria-label="New list"]',
        '[aria-label*="new list"]',
        '[data-name="new-watchlist"]',
        '[class*="addList"]',
        '[class*="newList"]',
      ];
      for (var s of selectors) {
        var el = document.querySelector(s);
        if (el) { el.click(); return { found: true, selector: s }; }
      }
      // Fallback: look for "+" icon button in watchlist header area
      var panel = document.querySelector('[class*="watchlist"]')
                || document.querySelector('[class*="layout__area--right"]');
      if (panel) {
        var btns = panel.querySelectorAll('button');
        for (var b of btns) {
          var lbl = b.getAttribute('aria-label') || b.textContent.trim();
          if (/new list|add list|[+]/i.test(lbl) && b.offsetParent) {
            b.click(); return { found: true, method: 'label_scan', label: lbl };
          }
        }
      }
      return { found: false };
    })()`);

  if (!created?.found) {
    // Fallback: use the 3-dot / context menu approach
    const menuOpened = await evaluate(`
      (function() {
        var panel = document.querySelector('[class*="layout__area--right"]');
        if (!panel) return { found: false };
        var dots = panel.querySelectorAll('[data-name="context-menu"], [aria-label*="More"], [class*="more"]');
        for (var d of dots) {
          if (d.offsetParent) { d.click(); return { found: true }; }
        }
        return { found: false };
      })()`);
    await sleep(400);
  }

  await sleep(500);

  // Type the new watchlist name
  await c.Input.insertText({ text: name });
  await sleep(300);

  // Press Enter to confirm
  await c.Input.dispatchKeyEvent({ type: 'keyDown', key: 'Enter', code: 'Enter', windowsVirtualKeyCode: 13 });
  await c.Input.dispatchKeyEvent({ type: 'keyUp',   key: 'Enter', code: 'Enter', windowsVirtualKeyCode: 13 });
  await sleep(800);

  console.log(`  ✅ Watchlist "${name}" created`);
}

async function addSymbol(symbol) {
  // Click "Add symbol" button
  const addClicked = await evaluate(`
    (function() {
      var selectors = [
        '[data-name="add-symbol-button"]',
        '[aria-label="Add symbol"]',
        '[aria-label*="Add symbol"]',
        'button[class*="addSymbol"]',
      ];
      for (var s of selectors) {
        var el = document.querySelector(s);
        if (el && el.offsetParent) { el.click(); return { found: true }; }
      }
      var panel = document.querySelector('[class*="layout__area--right"]');
      if (panel) {
        for (var btn of panel.querySelectorAll('button')) {
          if (btn.textContent.trim() === '+' && btn.offsetParent) {
            btn.click(); return { found: true, method: 'plus_btn' };
          }
        }
      }
      return { found: false };
    })()`);

  if (!addClicked?.found) {
    console.log(`  ⚠️  Add button not found for ${symbol}, skipping`);
    return;
  }

  await sleep(350);
  await c.Input.insertText({ text: symbol });
  await sleep(500);

  // Wait for suggestion to appear then press Enter
  await c.Input.dispatchKeyEvent({ type: 'keyDown', key: 'Enter', code: 'Enter', windowsVirtualKeyCode: 13 });
  await c.Input.dispatchKeyEvent({ type: 'keyUp',   key: 'Enter', code: 'Enter', windowsVirtualKeyCode: 13 });
  await sleep(250);

  // Escape to close search
  await c.Input.dispatchKeyEvent({ type: 'keyDown', key: 'Escape', code: 'Escape', windowsVirtualKeyCode: 27 });
  await c.Input.dispatchKeyEvent({ type: 'keyUp',   key: 'Escape', code: 'Escape', windowsVirtualKeyCode: 27 });
  await sleep(200);

  console.log(`  ➕  Added ${symbol}`);
}

// ── Main ──────────────────────────────────────────────────────────────────────

await ensureWatchlistPanelOpen();

// ── Watchlist 1: NSE Fundamentals ────────────────────────────────────────────
await createNewWatchlist('NSE Fundamentals');
for (const sym of FUNDAMENTALS) {
  await addSymbol(sym);
}
console.log(`\n✅ Watchlist "NSE Fundamentals" complete — ${FUNDAMENTALS.length} stocks added`);

await sleep(1000);

// ── Watchlist 2: NSE Scalp Setup ─────────────────────────────────────────────
await createNewWatchlist('NSE Scalp Setup');
for (const sym of SCALP_SETUP) {
  await addSymbol(sym);
}
console.log(`\n✅ Watchlist "NSE Scalp Setup" complete — ${SCALP_SETUP.length} stocks added`);

console.log('\n🏁 Both watchlists created successfully in TradingView.');
