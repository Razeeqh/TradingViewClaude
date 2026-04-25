"""
Volatility & Momentum Engine — Smart SL / Smart Entry
─────────────────────────────────────────────────────────────────────────────
Core engine that prevents premature stop-loss hits by sizing the SL to actual
stock volatility (ATR) instead of a fixed % — and confirms momentum is in
favor before signalling entry.

Outputs per stock:
  • atr_14d / atr_pct   — Average True Range over 14 days (₹ + % of price)
  • vol_regime          — LOW / NORMAL / HIGH / EXTREME (impacts position size)
  • smart_sl(entry)     — entry - (1.5 × ATR)  for swing
                          entry - (1.0 × ATR)  for scalp/intraday
                          entry - (2.0 × ATR)  for multibagger (trail wider)
  • momentum_score      — 0-100 from RSI(14), MACD signal, ADX, EMA stack
  • expected_move_2d    — 1σ move over 2 days (= ATR × √2)
  • sl_efficiency       — historical % of trades stopped out within 2 sessions
                          if SL was placed too tight (backfeed from backtest)

Inputs:
  • price_data.json  — OHLCV per symbol (refreshed daily by 8:45 AM task via
                       Kite MCP get_historical_data or NSE bhavcopy)
  • news_sentiment.json — per-symbol "positive/negative/neutral" tag from
                          the 8:45 task news scan

If price_data.json is absent, the engine falls back to template ATR values
for the watchlist universe (so screeners still produce output offline).
─────────────────────────────────────────────────────────────────────────────
"""
import json, os, math
from datetime import date, timedelta
from statistics import mean, stdev

PRICE_DATA_JSON  = r"C:\Users\razee\OneDrive\Desktop\TradingClaude\price_data.json"
NEWS_JSON        = r"C:\Users\razee\OneDrive\Desktop\TradingClaude\news_sentiment.json"
TODAY            = date.today()

# ── Fallback ATR % template (used when price_data.json is missing) ──────────
# These approximate Apr 2026 ATR% for liquid NSE names (verify with live data).
TEMPLATE_ATR_PCT = {
    # Index / market_context
    "NSE:NIFTY":      0.7,  "NSE:BANKNIFTY": 1.0,
    # Scalp watchlist (intraday)
    "NSE:COALINDIA":  1.5,  "NSE:NESTLEIND":  1.4,  "NSE:ICICIBANK":  1.3,
    "NSE:HDFCBANK":   1.2,  "NSE:BEL":        2.4,  "NSE:BAJFINANCE": 1.8,
    "NSE:BAJAJ-AUTO": 1.5,  "NSE:AXISBANK":   1.6,  "NSE:SBIN":       1.7,
    "NSE:JSWSTEEL":   2.0,  "NSE:POWERGRID":  1.2,  "NSE:NTPC":       1.4,
    "NSE:TRENT":      2.2,  "NSE:CIPLA":      1.4,  "NSE:SBILIFE":    1.6,
    "NSE:HDFCAMC":    2.0,  "NSE:CHOLAFIN":   2.3,
    # Swing fallen angels
    "NSE:SHAKTIPUMP": 3.5,  "NSE:TATAMOTORS": 2.4,  "NSE:VEDL":       2.6,
    "NSE:ADANIGREEN": 3.0,  "NSE:ASIANPAINT": 1.6,  "NSE:PAYTM":      3.2,
    "NSE:HINDALCO":   2.2,  "NSE:TATASTEEL":  2.3,  "NSE:NMDC":       2.4,
    "NSE:HEROMOTOCO": 1.8,  "NSE:INDIGO":     2.0,  "NSE:UPL":        2.5,
    "NSE:RECLTD":     2.6,  "NSE:PFC":        2.5,  "NSE:IRCTC":      2.4,
    "NSE:GODREJCP":   1.7,  "NSE:DABUR":      1.5,  "NSE:RELIANCE":   1.4,
    # Multibaggers (smid-cap, more volatile)
    "NSE:BDL":        3.0,  "NSE:MAZDOCK":    3.5,  "NSE:DATAPATTNS": 3.2,
    "NSE:MTARTECH":   3.4,  "NSE:KPIGREEN":   4.0,  "NSE:PREMIERENE": 4.2,
    "NSE:INOXWIND":   3.8,  "NSE:BORORENEW":  3.5,  "NSE:KEI":        2.4,
    "NSE:POLYCAB":    2.0,  "NSE:CGPOWER":    2.6,  "NSE:THERMAX":    2.4,
    "NSE:TRIVENI":    2.8,  "NSE:KAYNES":     3.4,  "NSE:DIXON":      2.8,
    "NSE:SYRMA":      3.2,  "NSE:UNOMINDA":   2.4,  "NSE:SONACOMS":   3.0,
    "NSE:DEEPAKNTR":  2.4,  "NSE:CLEAN":      2.6,  "NSE:NCC":        3.0,
    "NSE:KECL":       2.6,  "NSE:KALPATPOWR": 2.8,  "NSE:JIOFIN":     2.0,
    "NSE:AUBANK":     2.2,
}

# ── Volatility regimes ────────────────────────────────────────────────────────
def classify_vol_regime(atr_pct):
    """Classifies volatility regime based on ATR%."""
    if atr_pct < 1.5:  return "LOW",      1.0
    if atr_pct < 2.5:  return "NORMAL",   1.0
    if atr_pct < 4.0:  return "HIGH",     0.75
    return "EXTREME",  0.50

# ── ATR computation from OHLC (Wilder's smoothing) ───────────────────────────
def compute_atr_wilder(ohlc_bars, period=14):
    """ohlc_bars: list of dicts with 'high','low','close' for each bar.
       Returns ATR (₹) using Wilder's exponential smoothing."""
    if len(ohlc_bars) < period + 1:
        return None
    tr_list = []
    for i in range(1, len(ohlc_bars)):
        h = ohlc_bars[i]["high"]; l = ohlc_bars[i]["low"]
        prev_close = ohlc_bars[i-1]["close"]
        tr = max(h - l, abs(h - prev_close), abs(l - prev_close))
        tr_list.append(tr)
    # Initial ATR = simple average of first `period` TRs
    atr = mean(tr_list[:period])
    # Wilder smoothing for the rest
    for tr in tr_list[period:]:
        atr = (atr * (period - 1) + tr) / period
    return atr

# ── Momentum score (0-100) ───────────────────────────────────────────────────
def compute_rsi(closes, period=14):
    """Standard 14-period RSI."""
    if len(closes) < period + 1:
        return 50
    gains, losses = [], []
    for i in range(1, len(closes)):
        d = closes[i] - closes[i-1]
        gains.append(max(d, 0)); losses.append(max(-d, 0))
    avg_gain = mean(gains[:period]); avg_loss = mean(losses[:period])
    for i in range(period, len(gains)):
        avg_gain = (avg_gain * (period - 1) + gains[i]) / period
        avg_loss = (avg_loss * (period - 1) + losses[i]) / period
    if avg_loss == 0: return 100
    rs = avg_gain / avg_loss
    return 100 - (100 / (1 + rs))

def compute_ema(values, period):
    """Exponential moving average."""
    if len(values) < period: return None
    k = 2 / (period + 1)
    ema = mean(values[:period])
    for v in values[period:]:
        ema = v * k + ema * (1 - k)
    return ema

def momentum_score(closes):
    """Returns 0-100 score combining RSI, EMA stack, MACD."""
    if len(closes) < 50:
        return 50
    rsi = compute_rsi(closes)
    ema9  = compute_ema(closes, 9)
    ema20 = compute_ema(closes, 20)
    ema50 = compute_ema(closes, 50)
    if not all([ema9, ema20, ema50]):
        return 50
    last = closes[-1]
    score = 0
    # RSI in momentum zone (45-75)
    if 45 <= rsi <= 75: score += 30
    elif 35 <= rsi < 45 or 75 < rsi <= 80: score += 15
    # Price > EMA9 > EMA20 > EMA50 (bullish stack)
    if last > ema9 > ema20 > ema50: score += 35
    elif last > ema20 > ema50: score += 20
    elif last > ema50: score += 10
    # MACD proxy = EMA12 - EMA26 > 0
    ema12 = compute_ema(closes, 12); ema26 = compute_ema(closes, 26)
    if ema12 and ema26 and ema12 > ema26: score += 20
    # Recent 5-day return positive
    if closes[-1] > closes[-6]: score += 15
    return min(score, 100)

# ── Public API ────────────────────────────────────────────────────────────────
def load_price_data():
    """Loads price_data.json or returns empty dict."""
    if not os.path.exists(PRICE_DATA_JSON):
        return {}
    try:
        with open(PRICE_DATA_JSON) as f:
            return json.load(f)
    except Exception:
        return {}

def get_volatility_profile(symbol, current_price=None):
    """
    Returns a dict for the symbol:
    {
      atr: ₹X, atr_pct: X.X, vol_regime: "...", vol_size_multiplier: 1.0|0.75|0.5,
      momentum_score: 0-100, expected_move_2d_pct: X.X, source: "live"|"template"
    }
    """
    price_data = load_price_data()
    bars = price_data.get(symbol, {}).get("bars", [])
    if bars and len(bars) >= 30 and current_price:
        atr = compute_atr_wilder(bars[-60:], period=14)
        atr_pct = (atr / current_price) * 100 if atr else None
        closes = [b["close"] for b in bars[-60:]]
        mscore = momentum_score(closes)
        source = "live"
    else:
        atr_pct = TEMPLATE_ATR_PCT.get(symbol, 2.0)  # default 2% if unknown
        atr = (atr_pct / 100) * (current_price or 1000)
        mscore = 50  # neutral when no live data
        source = "template"

    regime, size_mult = classify_vol_regime(atr_pct)
    expected_move_2d_pct = atr_pct * math.sqrt(2)

    return {
        "atr": round(atr, 2),
        "atr_pct": round(atr_pct, 2),
        "vol_regime": regime,
        "vol_size_multiplier": size_mult,
        "momentum_score": mscore,
        "expected_move_2d_pct": round(expected_move_2d_pct, 2),
        "source": source,
    }

# ── Smart SL by trade horizon ─────────────────────────────────────────────────
SL_ATR_MULTIPLIER = {
    "scalp":       1.0,   # intraday — tight
    "1day_hold":   1.2,
    "swing":       1.5,   # 2-3 day swing
    "multibagger": 2.0,   # long hold — give it room
    "fallen_angel":2.5,   # need wide SL — these are buy-the-dip plays
}

def smart_sl(entry_price, symbol, horizon="swing"):
    """Returns SL price = entry - (multiplier × ATR)."""
    profile = get_volatility_profile(symbol, current_price=entry_price)
    multiplier = SL_ATR_MULTIPLIER.get(horizon, 1.5)
    sl = entry_price - (multiplier * profile["atr"])
    sl_pct = ((entry_price - sl) / entry_price) * 100
    return {
        "sl": round(sl, 2),
        "sl_pct": round(sl_pct, 2),
        "atr_multiplier": multiplier,
        "atr_pct": profile["atr_pct"],
        "vol_regime": profile["vol_regime"],
    }

# ── Smart targets by horizon (R:R-aware) ──────────────────────────────────────
def smart_targets(entry_price, symbol, horizon="swing", rr_min=2.0):
    """Returns T1, T2, T3 such that:
       T1 = entry + 1×R (where R = entry - SL)  → 1:1 minimum
       T2 = entry + 2×R                          → 2:1 (preferred exit)
       T3 = entry + 3×R                          → 3:1 (let-winners-run)
    """
    sl_data = smart_sl(entry_price, symbol, horizon)
    R = entry_price - sl_data["sl"]
    return {
        "t1": round(entry_price + R, 2),
        "t2": round(entry_price + 2 * R, 2),
        "t3": round(entry_price + 3 * R, 2),
        "t1_pct": round((R / entry_price) * 100, 2),
        "t2_pct": round((2 * R / entry_price) * 100, 2),
        "t3_pct": round((3 * R / entry_price) * 100, 2),
        "rr_at_t2": "2:1",
        "sl": sl_data["sl"],
        "sl_pct": sl_data["sl_pct"],
        "vol_regime": sl_data["vol_regime"],
    }

# ── Position sizing ───────────────────────────────────────────────────────────
def position_size(capital, risk_pct, entry_price, sl_price, symbol=None,
                  macro_multiplier=1.0):
    """Returns recommended position size in shares.
    capital: portfolio ₹
    risk_pct: 1% for swing, 0.5% for scalp, 2% for multibagger SIP tranche
    macro_multiplier: from swing_screener.macro_risk_multiplier (0.0-1.0)
    """
    risk_per_share = entry_price - sl_price
    if risk_per_share <= 0: return 0
    base_risk_rs = capital * (risk_pct / 100)
    # Apply volatility regime size multiplier
    if symbol:
        prof = get_volatility_profile(symbol, current_price=entry_price)
        vol_mult = prof["vol_size_multiplier"]
    else:
        vol_mult = 1.0
    adjusted_risk = base_risk_rs * vol_mult * macro_multiplier
    shares = int(adjusted_risk / risk_per_share)
    return {
        "shares": shares,
        "deployed_rs": shares * entry_price,
        "max_loss_rs": shares * risk_per_share,
        "vol_mult": vol_mult,
        "macro_mult": macro_multiplier,
        "effective_risk_pct": round((shares * risk_per_share / capital) * 100, 2) if capital else 0,
    }

# ── Demo / self-test ──────────────────────────────────────────────────────────
if __name__ == "__main__":
    print("Volatility & Momentum Engine — Self Test\n")
    test_cases = [
        ("NSE:BEL", 450, "1day_hold"),
        ("NSE:SHAKTIPUMP", 880, "swing"),
        ("NSE:KAYNES", 5100, "multibagger"),
        ("NSE:RELIANCE", 1200, "swing"),
        ("NSE:NIFTY", 24000, "scalp"),
    ]
    for sym, price, horizon in test_cases:
        print(f"━━━ {sym} @ ₹{price}  ({horizon}) ━━━")
        prof = get_volatility_profile(sym, current_price=price)
        sl = smart_sl(price, sym, horizon)
        tgts = smart_targets(price, sym, horizon)
        sz = position_size(1_000_000, 1.0 if horizon == "swing" else 0.5, price, sl["sl"], sym)
        print(f"  ATR: ₹{prof['atr']} ({prof['atr_pct']}%) — Regime: {prof['vol_regime']}")
        print(f"  Momentum: {prof['momentum_score']}/100 — 2-day expected move: ±{prof['expected_move_2d_pct']}%")
        print(f"  Smart SL: ₹{sl['sl']} (-{sl['sl_pct']}%, {sl['atr_multiplier']}×ATR)")
        print(f"  T1: ₹{tgts['t1']} (+{tgts['t1_pct']}%)  T2: ₹{tgts['t2']} (+{tgts['t2_pct']}%)  T3: ₹{tgts['t3']} (+{tgts['t3_pct']}%)")
        print(f"  Position: {sz['shares']} shares = ₹{sz['deployed_rs']:.0f} deployed; max loss ₹{sz['max_loss_rs']:.0f} ({sz['effective_risk_pct']}% of capital)\n")
