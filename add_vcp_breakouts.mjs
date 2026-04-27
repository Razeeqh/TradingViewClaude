/**
 * Adds the FULL 3-tier VCP universe (Largecap + Midcap + Smallcap) to
 * TradingView watchlist — exactly as stockexploder methodology recommends.
 *
 * Risk-balanced: largecaps for steadier 5-12% moves, midcaps for 10-18%,
 * smallcaps for explosive 15-30%.
 *
 * Order: priority by stage (BREAKING OUT → PIVOT-READY → CONTRACTING)
 * within each cap tier, so the watchlist ordering reflects what to act on.
 */
import { connect, evaluate, getClient } from './src/connection.js';

await connect();
const c = await getClient();
const sleep = ms => new Promise(r => setTimeout(r, ms));

// ── 3-TIER VCP UNIVERSE — 21 stocks ────────────────────────────────────────
const VCP_PICKS = [
  // ══════════════════════════════════════════════════════════════════════
  // ██ LARGECAP VCP (5)  — MCap > ₹50,000 cr  — moves 5-12% in 3-10 days
  // ══════════════════════════════════════════════════════════════════════
  'NSE:ADANIPORTS',   // 🟢 PIVOT — READY  (at 52w high ₹1,628 — breakout imminent)
  'NSE:BAJFINANCE',   // 🟢 PIVOT — READY  (tight base ₹900-925)
  'NSE:BHEL',         // 🟢 PIVOT — READY  (₹1.2L cr order book + capex super-cycle)
  'NSE:TATAPOWER',    // 🔵 CONTRACTING    (3 contractions, base tightening)
  'NSE:PREMIERENE',   // 🔵 CONTRACTING    (Solar — ALMM regime supportive)

  // ══════════════════════════════════════════════════════════════════════
  // ██ MIDCAP VCP (8)  — MCap ₹15k-50k cr  — moves 10-18% in 2-5 days
  // ══════════════════════════════════════════════════════════════════════
  'NSE:HBLENGINE',    // 🚀 BREAKING OUT   (vol 2.8x avg + 52w high broken)
  'NSE:DATAPATTNS',   // 🚀 BREAKING OUT   (broke ₹3,900, Nippon MF bulk buy)
  'NSE:AVALON',       // 🚀 BREAKING OUT   (vol 2.8x — 52w high broken)
  'NSE:KAYNES',       // 🔵 CONTRACTING    (semicon OSAT + ICICI Pru bulk)
  'NSE:JYOTICNC',     // 🔵 CONTRACTING    (defence CNC orders ramping)
  'NSE:ASTRAMICRO',   // 🔵 CONTRACTING    (radar exports + DRDO supply)
  'NSE:KIRLOSBROS',   // 🔵 CONTRACTING    (Jal Jeevan + naval pumps)
  'NSE:TITAGARH',     // 🔵 CONTRACTING    (Vande Bharat + freight wagons)

  // ══════════════════════════════════════════════════════════════════════
  // ██ SMALLCAP VCP (8)  — MCap < ₹15k cr  — moves 15-30% in 1-3 days
  // ══════════════════════════════════════════════════════════════════════
  'NSE:AZAD',         // 🚀 BREAKING OUT   (GE Aero + Rolls-Royce supplier)
  'NSE:SEDEMAC',      // 🟢 PIVOT — READY  (textbook 8→4→2% contractions, IPO Mar-26)
  'NSE:RPEL',         // 🟢 PIVOT — READY  (silica ramming mass — steel capex)
  'NSE:IDEAFORGE',    // 🔵 CONTRACTING    (defence drones, Hormuz tension)
  'NSE:PARAS',        // 🔵 CONTRACTING    (optronics + space)
  'NSE:ZAGGLE',       // 🔵 CONTRACTING    (B2B fintech SaaS)
  'NSE:CYIENTDLM',    // 🔵 CONTRACTING    (aerospace EMS — Q4 beat)
  'NSE:GANECOS',      // 🟡 BASING         (recycled PET — ESG mandate)
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
  if (!clicked?.found) { console.log(`  ⚠️  Skipped ${symbol} (button not found)`); return; }
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
console.log('\n🚀 Adding 3-tier VCP Breakout candidates to TradingView watchlist:\n');
console.log('   LARGECAP VCP (5)  — moves 5-12%  in 3-10 days');
console.log('   MIDCAP VCP   (8)  — moves 10-18% in 2-5 days');
console.log('   SMALLCAP VCP (8)  — moves 15-30% in 1-3 days');
console.log('\n   [Mark Minervini SEPA + stockexploder methodology]\n');

for (const sym of VCP_PICKS) await addSymbol(sym);

console.log(`\n🏁 Done — ${VCP_PICKS.length} VCP candidates added across 3 cap tiers.`);
console.log('\n📌 Trade discipline (every cap tier):');
console.log('   • Wait for pivot breakout + volume ≥ 2x avg before entry');
console.log('   • Hard SL just below pivot — no exceptions');
console.log('   • BOOK: T1 +5% → 40%   T2 +14% → 35%   T3 +24% → trail 25%');
console.log('   • Smaller size on smallcaps — same 1% portfolio risk per trade');
console.log('   • Max 2 VCP positions concurrent');
