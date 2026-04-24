import { connect, evaluate } from './src/connection.js';
import { getQuote } from './src/core/data.js';

await connect();

// Navigate chart to NSE:NIFTY for market context verification
await evaluate(`
  (function() {
    try {
      window.TradingViewApi._activeChartWidgetWV.value().setSymbol('NSE:NIFTY', function() {});
    } catch(e) {}
  })()`);

await new Promise(r => setTimeout(r, 5000));

try {
  const quote = await getQuote({ symbol: 'NSE:NIFTY' });
  console.log('NSE:NIFTY QUOTE:', JSON.stringify(quote, null, 2));
} catch(e) {
  console.log('Quote error:', e.message);
}
