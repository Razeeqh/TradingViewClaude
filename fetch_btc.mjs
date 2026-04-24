import { connect, evaluate } from './src/connection.js';

await connect();

// Fetch BTC price via Binance public API from within TradingView's browser context
const result = await evaluate(`
  new Promise((resolve) => {
    fetch('https://api.binance.com/api/v3/ticker/price?symbol=BTCUSDT')
      .then(r => r.json())
      .then(d => resolve(d))
      .catch(e => resolve({ error: e.message }));
  })`, { awaitPromise: true });

console.log('BTCUSDT price from Binance API:', JSON.stringify(result, null, 2));
