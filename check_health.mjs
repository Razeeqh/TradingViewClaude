import CDP from 'chrome-remote-interface';

const client = await CDP({ port: 9222 });
await client.Runtime.enable();

// Health check - which chart is active
const health = await client.Runtime.evaluate({
  expression: `(function() {
    try {
      var chart = window.TradingViewApi._activeChartWidgetWV.value();
      return JSON.stringify({ symbol: chart.symbol(), resolution: chart.resolution(), apiAvailable: true });
    } catch(e) {
      return JSON.stringify({ apiAvailable: false, error: e.message });
    }
  })()`,
  returnByValue: true
});
console.log('HEALTH:', health.result.value);

// Fetch BTC price via TradingView's quote data
const btcPrice = await client.Runtime.evaluate({
  expression: `(function() {
    try {
      var tvApi = window.TradingViewApi;
      return JSON.stringify({ attempted: true });
    } catch(e) {
      return JSON.stringify({ error: e.message });
    }
  })()`,
  returnByValue: true
});
console.log('BTC_CHECK:', btcPrice.result.value);

await client.close();
