/**
 * Adds today's TOP VCP breakout candidates to TradingView watchlist.
 * Source of truth: NSE_VCP_Breakouts.xlsx + vcp_fresh.json from daily 8 AM scan.
 *
 * VCP candidates target 8-20% intraday/short-swing moves vs 1-2% large-cap scalps.
 * Mark Minervini methodology + stockexploder reference.
 */
import { connect, evaluate, getClient } from './src/connection.js';

await connect();
const c = await getClient();
const sleep = ms => new Promise(r => setTimeout(r, ms));

// TOP VCP BREAKOUT CANDIDATES (refresh daily via 8 AM scheduled task)
// Priority order: 🚀 BREAKING OUT → 🟢 PIVOT — READY → 🔵 CONTRACTING (top conviction)
const VCP_TOP_PICKS = [
  // 🚀 BREAKING OUT — execute at open
  'NSE:RPEL',         // Raghav Productivity — VCP breakout confirmed, +20% expected 3D
  // 🟢 PIVOT — READY — enter on breakout confirmation
  'NSE:SEDEMAC',      // Sedemac Mechatronics — pivot ₹1,820, +18% expected
  // 🔵 CONTRACTING — top conviction, watch closely this week
  'NSE:AZAD',         // Azad Engineering — aerospace + defence, +16% expected
  'NSE:KAYNES',       // Kaynes Technology — EMS + semicon, +14% expected
  'NSE:ZAGGLE',       // Zaggle — fintech B2B, +18% expected
  'NSE:CYIENTDLM',    // Cyient DLM — aerospace EMS, +16% expected
  'NSE:IDEAFORGE',    // ideaForge — defence drones, +15% expected
  'NSE:DATAPATTNS',   // Data Patterns — defence electronics
  'NSE:ASTRAMICRO',   // Astra Microwave — defence radar
  'NSE:PARAS',        // Paras Defence — optronics + space
  'NSE:HBLENGINE',    // HBL Power — Kavach + submarine batteries
  'NSE:TITAGARH',     // Titagarh Rail
  'NSE:JYOTICNC',     // Jyoti CNC — machine tools
  'NSE:ADITYA-VISION',// Aditya Vision — retail
  'NSE:KIRLOSBROS',   // Kirloskar Brothers — pumps
  'NSE:TIPSINDLTD',   // Tips Industries — music
  'NSE:GANESHA',      // Ganesha Ecosphere — recycled PET
  'NSE:AVALON',       // Avalon Tech — EMS aero
  'NSE:PREMIERENE',   // Premier Energies — solar cells
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
console.log('\n🚀 Adding VCP Breakout candidates to TradingView watchlist:\n');
console.log('   (8-20% intraday move targets — Minervini VCP + stockexploder methodology)\n');

for (const sym of VCP_TOP_PICKS) await addSymbol(sym);

console.log(`\n🏁 Done — ${VCP_TOP_PICKS.length} VCP candidates added.`);
console.log('\n📌 Trade discipline:');
console.log('   • Wait for ACTUAL pivot breakout + volume ≥ 2x avg');
console.log('   • Hard SL 5% below pivot — no exceptions');
console.log('   • Book 30% at +5-8%, 40% at +12-15%, trail rest');
console.log('   • Max 2 VCP trades concurrent');
