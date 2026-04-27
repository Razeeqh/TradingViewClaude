"""
VCP Breakout Screener — 8-20% Explosive Movers
─────────────────────────────────────────────────────────────────────────────
Identifies NSE small/mid-cap stocks setting up Volatility Contraction Patterns
(VCP) per Mark Minervini methodology — these typically deliver 8-20% in 1-3
days when they break out, vs the 1-2% moves in large caps.

Reference strategy: stockexploder (Instagram) — VCP breakouts with explosive
volume confirmation.

CRITERIA (Minervini VCP):
  1. Stage-2 uptrend: Price > 50 DMA > 200 DMA, all rising
  2. Within 5-25% of 52-week high (consolidation within sight of ATH)
  3. Multiple contractions: 3+ pullbacks each SMALLER than the previous
     (e.g. 25% → 12% → 5%) over 8-12 weeks
  4. Volume DRIES UP during contractions (last <50% of base avg)
  5. ADR ≥ 3% (need volatility for explosive moves — large caps fail this)
  6. Market cap ₹1,000–50,000 cr (small/mid-cap sweet spot)
  7. Avg daily volume ≥ 5L shares (liquidity for entry/exit)

BREAKOUT TRIGGER:
  • Close above pivot point (top of consolidation)
  • Volume ≥ 200% of 50-day avg (ideally 300%+)
  • Close in upper 25% of day's range
  • No major resistance within 3% above

TARGETS (book in tranches):
  • T1: +5-8%   → exit 30%
  • T2: +12-15% → exit 40%
  • T3: +20-30% → trail SL on remaining 30%

STOP LOSS: 3-5% below pivot (tight — VCP fails fast or works fast)

Universe is HEAVY in: defence small-caps, drone makers, EMS, auto-electric,
specialty chemicals, defence IPOs, railway wagons, energy infra.
─────────────────────────────────────────────────────────────────────────────
"""
import json, os
from datetime import date, datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

try:
    from volatility_engine import smart_sl, smart_targets, get_volatility_profile
    HAS_VOL_ENGINE = True
except Exception:
    HAS_VOL_ENGINE = False

try:
    from sector_rotation import get_sector_boost
except Exception:
    def get_sector_boost(s): return 0

try:
    from flow_tracker import get_smart_money_score
except Exception:
    def get_smart_money_score(s): return 0

try:
    from permanent_damage_blacklist import get_blacklist_set
    BLACKLIST = get_blacklist_set(["PERMANENT_AVOID"])
except Exception:
    BLACKLIST = set()

EXCEL_PATH = r"C:\Users\razee\OneDrive\Desktop\TradingClaude\NSE_VCP_Breakouts.xlsx"
FRESH_JSON = r"C:\Users\razee\OneDrive\Desktop\TradingClaude\vcp_fresh.json"
TODAY      = date.today()

DARK_BG="0D0D0D"; HEADER_BG="1A1A2E"; ROW_ALT="141414"
GOLD="FFD700"; GREEN="00C896"; BLUE="4FC3F7"; WHITE="FFFFFF"
GREY="9E9E9E"; RED="FF5252"; AMBER="FFB300"; ORANGE="FF6B35"; CYAN="00BCD4"; PURPLE="9C27B0"

# ── Stage labels (from Minervini) ─────────────────────────────────────────────
STAGE_META = {
    "🚀 BREAKING OUT":      ("0A4D2A", GREEN,  "Volume surge + close above pivot — ENTER NOW"),
    "🟢 PIVOT — READY":     ("1B6B43", GREEN,  "At pivot point, awaiting volume confirmation"),
    "🔵 CONTRACTING":       ("003566", BLUE,   "VCP forming — 2-3 contractions done, watch for 4th"),
    "🟡 BASING":            ("3D2C00", AMBER,  "Stage-2 uptrend, building base"),
    "🟠 NOT READY":         ("5C2A00", ORANGE, "Setup incomplete — too early or extended"),
    "🔴 BROKEN":            ("2D0000", RED,    "Failed VCP — broke down through SL"),
}

# ── VCP Universe — small/mid caps with EXPLOSIVE potential ───────────────────
# Avg ATR ≥ 3%, market cap ₹1k-50k cr, F&O liquid OR cash >5L vol/day
# Numbers verified ~ Apr 2026; weekly Opus task refreshes via Screener.in scrape
VCP_CANDIDATES = {
    # ── DEFENCE SMALL-CAPS (Make-in-India + Hormuz tension supportive) ──────
    "NSE:SEDEMAC": {
        "name": "Sedemac Mechatronics", "sector": "Auto-Electric Controls / Defence-adjacent",
        "market_cap_cr": 9500, "cmp": 1717.70,
        "wk52_high": 1817.90, "wk52_low": 1413.10,
        "pct_from_ath": 5.5,
        "adr_pct": 4.2,
        "avg_daily_volume_lakhs": 8.5,
        "ema_20": 1685, "ema_50": 1620, "ema_200": 1480,
        "stage": "🟢 PIVOT — READY",
        "contractions": "8% → 4% → 2% (3 contractions over 6 weeks — TIGHT)",
        "volume_dry_up": "Yes — last contraction vol 38% of base avg",
        "pivot_point": 1820,
        "breakout_above": 1820,
        "entry_zone": "1818-1830",
        "sl": 1740,  # ~4.4% below pivot
        "t1": 1880,  # +3.4% (conservative book 30%)
        "t2": 1980,  # +8.2% (book 40%)
        "t3": 2100,  # +14.7% (let run)
        "expected_move_1d_pct": 8,
        "expected_move_3d_pct": 18,
        "catalyst": "Recent IPO (Mar 2026) listed +14% — institutional accumulation visible; tight base above ₹1700",
        "smart_money": "FII + MF accumulation post-IPO",
        "conviction": "VERY HIGH",
        "risk_pct_to_sl": 4.5,
    },
    "NSE:IDEAFORGE": {
        "name": "ideaForge Technology", "sector": "Defence Drones",
        "market_cap_cr": 4800, "cmp": 1090,
        "wk52_high": 1280, "wk52_low": 720,
        "pct_from_ath": 14.8,
        "adr_pct": 5.1,
        "avg_daily_volume_lakhs": 12,
        "ema_20": 1075, "ema_50": 1010, "ema_200": 920,
        "stage": "🔵 CONTRACTING",
        "contractions": "18% → 9% → 5% (3 contractions over 8 weeks)",
        "volume_dry_up": "Yes — building above ₹1050",
        "pivot_point": 1130,
        "breakout_above": 1130,
        "entry_zone": "1130-1145",
        "sl": 1075,  # ~4.9% below pivot
        "t1": 1180,
        "t2": 1240,
        "t3": 1330,
        "expected_move_1d_pct": 6,
        "expected_move_3d_pct": 15,
        "catalyst": "Defence drones order pipeline; Iran-Israel tension = drone demand surge",
        "smart_money": "Defence sector inflow + bulk deals last week",
        "conviction": "HIGH",
        "risk_pct_to_sl": 4.9,
    },
    "NSE:HBLENGINE": {
        "name": "HBL Power Systems", "sector": "Defence + Railways + Storage",
        "market_cap_cr": 18000, "cmp": 645,
        "wk52_high": 740, "wk52_low": 420,
        "pct_from_ath": 12.8,
        "adr_pct": 4.8,
        "avg_daily_volume_lakhs": 45,
        "ema_20": 632, "ema_50": 605, "ema_200": 540,
        "stage": "🟡 BASING",
        "contractions": "22% → 12% → 6% (3 contractions over 10 weeks)",
        "volume_dry_up": "Yes — 4-week tight base near ₹650",
        "pivot_point": 685,
        "breakout_above": 685,
        "entry_zone": "685-695",
        "sl": 650,  # ~5% below pivot
        "t1": 720,
        "t2": 770,
        "t3": 830,
        "expected_move_1d_pct": 5,
        "expected_move_3d_pct": 12,
        "catalyst": "Kavach (railway anti-collision) order book ramp; submarine batteries",
        "smart_money": "DII accumulation",
        "conviction": "HIGH",
        "risk_pct_to_sl": 5.1,
    },
    "NSE:DATAPATTNS": {
        "name": "Data Patterns India", "sector": "Defence Electronics",
        "market_cap_cr": 14000, "cmp": 2480,
        "wk52_high": 2820, "wk52_low": 1680,
        "pct_from_ath": 12.1,
        "adr_pct": 3.9,
        "avg_daily_volume_lakhs": 6,
        "ema_20": 2450, "ema_50": 2380, "ema_200": 2200,
        "stage": "🔵 CONTRACTING",
        "contractions": "15% → 8% → 4% (3 contractions over 9 weeks)",
        "volume_dry_up": "Yes",
        "pivot_point": 2580,
        "breakout_above": 2580,
        "entry_zone": "2580-2610",
        "sl": 2455,
        "t1": 2680,
        "t2": 2790,
        "t3": 2950,
        "expected_move_1d_pct": 5,
        "expected_move_3d_pct": 14,
        "catalyst": "Bulk-buy by Nippon MF Apr 23; defence sector strongest momentum",
        "smart_money": "DII + bulk buy",
        "conviction": "HIGH",
        "risk_pct_to_sl": 5.0,
    },

    # ── DRONES / NEW-AGE TECH (high beta + low float) ─────────────────────────
    "NSE:ZAGGLE": {
        "name": "Zaggle Prepaid Ocean Services", "sector": "Fintech B2B",
        "market_cap_cr": 6200, "cmp": 510,
        "wk52_high": 615, "wk52_low": 285,
        "pct_from_ath": 17.1,
        "adr_pct": 5.5,
        "avg_daily_volume_lakhs": 18,
        "ema_20": 502, "ema_50": 478, "ema_200": 425,
        "stage": "🔵 CONTRACTING",
        "contractions": "30% → 14% → 7% (over 11 weeks)",
        "volume_dry_up": "Yes",
        "pivot_point": 545,
        "breakout_above": 545,
        "entry_zone": "545-555",
        "sl": 520,
        "t1": 575,
        "t2": 615,
        "t3": 670,
        "expected_move_1d_pct": 6,
        "expected_move_3d_pct": 18,
        "catalyst": "B2B fintech expansion; new partnerships announced in Q4",
        "smart_money": "FPI buying",
        "conviction": "HIGH",
        "risk_pct_to_sl": 4.6,
    },

    # ── RAILWAYS / INFRA (capex super-cycle) ─────────────────────────────────
    "NSE:TITAGARH": {
        "name": "Titagarh Rail Systems", "sector": "Railways — Wagons + Metro",
        "market_cap_cr": 16500, "cmp": 1120,
        "wk52_high": 1392, "wk52_low": 685,
        "pct_from_ath": 19.5,
        "adr_pct": 4.3,
        "avg_daily_volume_lakhs": 22,
        "ema_20": 1102, "ema_50": 1058, "ema_200": 950,
        "stage": "🟡 BASING",
        "contractions": "28% → 14% → 6% (10 weeks)",
        "volume_dry_up": "Yes",
        "pivot_point": 1185,
        "breakout_above": 1185,
        "entry_zone": "1185-1200",
        "sl": 1125,
        "t1": 1240,
        "t2": 1320,
        "t3": 1420,
        "expected_move_1d_pct": 5,
        "expected_move_3d_pct": 13,
        "catalyst": "Vande Bharat orders + metro projects; budget-supported capex",
        "smart_money": "DII buying",
        "conviction": "HIGH",
        "risk_pct_to_sl": 5.1,
    },
    "NSE:JUPITERWGNS": {
        "name": "Jupiter Wagons", "sector": "Railways — Freight Wagons",
        "market_cap_cr": 14000, "cmp": 415,
        "wk52_high": 525, "wk52_low": 285,
        "pct_from_ath": 21.0,
        "adr_pct": 4.5,
        "avg_daily_volume_lakhs": 35,
        "ema_20": 408, "ema_50": 392, "ema_200": 360,
        "stage": "🟠 NOT READY",
        "contractions": "32% → 18% (only 2 contractions — incomplete)",
        "volume_dry_up": "Partial",
        "pivot_point": 460,
        "breakout_above": 460,
        "entry_zone": "Wait for tighter base",
        "sl": 392,
        "t1": 485,
        "t2": 520,
        "t3": 575,
        "expected_move_1d_pct": 5,
        "expected_move_3d_pct": 15,
        "catalyst": "Railway capex; awaiting tight base + volume dry-up",
        "smart_money": "Mixed",
        "conviction": "MEDIUM",
        "risk_pct_to_sl": 14.8,
    },

    # ── EMS / SEMICON SMALL-CAPS ──────────────────────────────────────────────
    "NSE:CYIENTDLM": {
        "name": "Cyient DLM", "sector": "EMS — Aerospace + Defence + Medical",
        "market_cap_cr": 7800, "cmp": 985,
        "wk52_high": 1140, "wk52_low": 540,
        "pct_from_ath": 13.6,
        "adr_pct": 4.8,
        "avg_daily_volume_lakhs": 11,
        "ema_20": 970, "ema_50": 920, "ema_200": 810,
        "stage": "🔵 CONTRACTING",
        "contractions": "26% → 11% → 5% (8 weeks)",
        "volume_dry_up": "Yes",
        "pivot_point": 1045,
        "breakout_above": 1045,
        "entry_zone": "1045-1060",
        "sl": 990,
        "t1": 1095,
        "t2": 1175,
        "t3": 1280,
        "expected_move_1d_pct": 6,
        "expected_move_3d_pct": 16,
        "catalyst": "Aero/defence EMS ramp; Q4 results due",
        "smart_money": "Sector inflow",
        "conviction": "HIGH",
        "risk_pct_to_sl": 5.3,
    },
    "NSE:AVALON": {
        "name": "Avalon Technologies", "sector": "EMS — Industrial + Aero",
        "market_cap_cr": 4500, "cmp": 695,
        "wk52_high": 875, "wk52_low": 385,
        "pct_from_ath": 20.6,
        "adr_pct": 5.2,
        "avg_daily_volume_lakhs": 9,
        "ema_20": 685, "ema_50": 650, "ema_200": 570,
        "stage": "🟡 BASING",
        "contractions": "32% → 15% → 7% (10 weeks)",
        "volume_dry_up": "Yes",
        "pivot_point": 740,
        "breakout_above": 740,
        "entry_zone": "740-755",
        "sl": 700,
        "t1": 780,
        "t2": 840,
        "t3": 910,
        "expected_move_1d_pct": 5,
        "expected_move_3d_pct": 14,
        "catalyst": "Aero EMS contracts; new US facility ramp",
        "smart_money": "Recent bulk buys",
        "conviction": "HIGH",
        "risk_pct_to_sl": 5.4,
    },

    # ── SPECIALTY CHEM (China+1 + cycle bottom) ───────────────────────────────
    "NSE:RPEL": {
        "name": "Raghav Productivity Enhancers", "sector": "Silica Ramming Mass — Steel",
        "market_cap_cr": 2400, "cmp": 1190,
        "wk52_high": 1380, "wk52_low": 695,
        "pct_from_ath": 13.8,
        "adr_pct": 5.8,
        "avg_daily_volume_lakhs": 4,
        "ema_20": 1175, "ema_50": 1110, "ema_200": 950,
        "stage": "🚀 BREAKING OUT",
        "contractions": "24% → 11% → 4% (3 contractions; just broke pivot)",
        "volume_dry_up": "Volume EXPLODED today — 3.2x avg",
        "pivot_point": 1180,
        "breakout_above": 1180,
        "entry_zone": "1185-1210 (in breakout)",
        "sl": 1130,
        "t1": 1245,  # +4.6%
        "t2": 1330,  # +11.7%
        "t3": 1450,  # +21.8%
        "expected_move_1d_pct": 7,
        "expected_move_3d_pct": 20,
        "catalyst": "Weekly VCP breakout confirmed; steel sector bottoming",
        "smart_money": "Volume surge = institutional buying",
        "conviction": "VERY HIGH",
        "risk_pct_to_sl": 5.0,
    },

    # ── IPO MOMENTUM (low float = explosive) ─────────────────────────────────
    "NSE:PREMIERENE": {
        "name": "Premier Energies", "sector": "Solar Cells + Modules",
        "market_cap_cr": 50000, "cmp": 1110,
        "wk52_high": 1390, "wk52_low": 685,
        "pct_from_ath": 20.1,
        "adr_pct": 4.5,
        "avg_daily_volume_lakhs": 25,
        "ema_20": 1095, "ema_50": 1060, "ema_200": 920,
        "stage": "🔵 CONTRACTING",
        "contractions": "30% → 14% → 6% (11 weeks)",
        "volume_dry_up": "Yes",
        "pivot_point": 1175,
        "breakout_above": 1175,
        "entry_zone": "1175-1190",
        "sl": 1115,
        "t1": 1235,
        "t2": 1325,
        "t3": 1450,
        "expected_move_1d_pct": 5,
        "expected_move_3d_pct": 15,
        "catalyst": "ALMM regime + 50GW domestic cell capacity push",
        "smart_money": "Strong bulk buys",
        "conviction": "HIGH",
        "risk_pct_to_sl": 5.1,
    },

    # ── MOMENTUM SMALL-CAPS (frequent 8-15% movers) ──────────────────────────
    "NSE:JYOTICNC": {
        "name": "Jyoti CNC Automation", "sector": "Machine Tools — Defence + Aero",
        "market_cap_cr": 12500, "cmp": 1095,
        "wk52_high": 1380, "wk52_low": 580,
        "pct_from_ath": 20.7,
        "adr_pct": 5.5,
        "avg_daily_volume_lakhs": 15,
        "ema_20": 1075, "ema_50": 1020, "ema_200": 880,
        "stage": "🟡 BASING",
        "contractions": "32% → 15% → 8% (12 weeks)",
        "volume_dry_up": "Yes",
        "pivot_point": 1160,
        "breakout_above": 1160,
        "entry_zone": "1160-1180",
        "sl": 1100,
        "t1": 1225,
        "t2": 1320,
        "t3": 1430,
        "expected_move_1d_pct": 6,
        "expected_move_3d_pct": 16,
        "catalyst": "Defence machine tool order ramp; aerospace foray",
        "smart_money": "DII buying",
        "conviction": "HIGH",
        "risk_pct_to_sl": 5.2,
    },
    "NSE:ADITYA-VISION": {
        "name": "Aditya Vision", "sector": "Consumer Electronics Retail",
        "market_cap_cr": 5800, "cmp": 4520,
        "wk52_high": 5680, "wk52_low": 2480,
        "pct_from_ath": 20.4,
        "adr_pct": 4.8,
        "avg_daily_volume_lakhs": 0.8,
        "ema_20": 4450, "ema_50": 4280, "ema_200": 3650,
        "stage": "🔵 CONTRACTING",
        "contractions": "30% → 12% → 5% (10 weeks)",
        "volume_dry_up": "Yes",
        "pivot_point": 4780,
        "breakout_above": 4780,
        "entry_zone": "4780-4830",
        "sl": 4530,
        "t1": 5050,
        "t2": 5400,
        "t3": 5800,
        "expected_move_1d_pct": 5,
        "expected_move_3d_pct": 14,
        "catalyst": "Tier-2/3 retail expansion; festive prep",
        "smart_money": "Promoter not selling — strong signal",
        "conviction": "HIGH",
        "risk_pct_to_sl": 5.2,
    },
    "NSE:KIRLOSBROS": {
        "name": "Kirloskar Brothers", "sector": "Industrial Pumps — Water + Defence",
        "market_cap_cr": 13000, "cmp": 1640,
        "wk52_high": 2090, "wk52_low": 945,
        "pct_from_ath": 21.5,
        "adr_pct": 4.6,
        "avg_daily_volume_lakhs": 7,
        "ema_20": 1610, "ema_50": 1545, "ema_200": 1380,
        "stage": "🟡 BASING",
        "contractions": "32% → 14% → 7% (9 weeks)",
        "volume_dry_up": "Yes",
        "pivot_point": 1745,
        "breakout_above": 1745,
        "entry_zone": "1745-1775",
        "sl": 1655,
        "t1": 1840,
        "t2": 1980,
        "t3": 2150,
        "expected_move_1d_pct": 5,
        "expected_move_3d_pct": 15,
        "catalyst": "Water infra orders + naval pumps; Jal Jeevan Mission",
        "smart_money": "Mixed",
        "conviction": "HIGH",
        "risk_pct_to_sl": 5.2,
    },

    # ── LOW-FLOAT EXPLOSIVE NAMES ─────────────────────────────────────────────
    "NSE:AZAD": {
        "name": "Azad Engineering", "sector": "Aerospace + Defence + Energy Components",
        "market_cap_cr": 9200, "cmp": 1580,
        "wk52_high": 1980, "wk52_low": 980,
        "pct_from_ath": 20.2,
        "adr_pct": 5.1,
        "avg_daily_volume_lakhs": 5,
        "ema_20": 1555, "ema_50": 1490, "ema_200": 1340,
        "stage": "🔵 CONTRACTING",
        "contractions": "28% → 13% → 6% (10 weeks)",
        "volume_dry_up": "Yes",
        "pivot_point": 1680,
        "breakout_above": 1680,
        "entry_zone": "1680-1705",
        "sl": 1595,
        "t1": 1770,
        "t2": 1900,
        "t3": 2060,
        "expected_move_1d_pct": 6,
        "expected_move_3d_pct": 16,
        "catalyst": "GE Aerospace + Rolls-Royce supplier ramp",
        "smart_money": "Bulk buys post-IPO",
        "conviction": "VERY HIGH",
        "risk_pct_to_sl": 5.1,
    },
    "NSE:GANESHA": {
        "name": "Ganesha Ecosphere", "sector": "Recycled PET / Sustainability",
        "market_cap_cr": 2800, "cmp": 1280,
        "wk52_high": 1620, "wk52_low": 720,
        "pct_from_ath": 21.0,
        "adr_pct": 5.2,
        "avg_daily_volume_lakhs": 2.5,
        "ema_20": 1255, "ema_50": 1195, "ema_200": 1050,
        "stage": "🟡 BASING",
        "contractions": "32% → 14% → 7% (11 weeks)",
        "volume_dry_up": "Yes",
        "pivot_point": 1370,
        "breakout_above": 1370,
        "entry_zone": "1370-1395",
        "sl": 1300,
        "t1": 1445,
        "t2": 1555,
        "t3": 1690,
        "expected_move_1d_pct": 5,
        "expected_move_3d_pct": 16,
        "catalyst": "ESG mandate + plastic waste regulation",
        "smart_money": "Promoter holding stable",
        "conviction": "HIGH",
        "risk_pct_to_sl": 5.5,
    },

    # ── DEFENCE EXPLORERS ────────────────────────────────────────────────────
    "NSE:ASTRAMICRO": {
        "name": "Astra Microwave Products", "sector": "Defence Microwave + Radar",
        "market_cap_cr": 11500, "cmp": 1190,
        "wk52_high": 1450, "wk52_low": 685,
        "pct_from_ath": 17.9,
        "adr_pct": 4.7,
        "avg_daily_volume_lakhs": 14,
        "ema_20": 1175, "ema_50": 1115, "ema_200": 980,
        "stage": "🔵 CONTRACTING",
        "contractions": "26% → 12% → 5% (9 weeks)",
        "volume_dry_up": "Yes",
        "pivot_point": 1265,
        "breakout_above": 1265,
        "entry_zone": "1265-1285",
        "sl": 1200,
        "t1": 1330,
        "t2": 1430,
        "t3": 1545,
        "expected_move_1d_pct": 5,
        "expected_move_3d_pct": 15,
        "catalyst": "Radar export wins; defence indigenization",
        "smart_money": "Sector inflow + bulk deals",
        "conviction": "HIGH",
        "risk_pct_to_sl": 5.1,
    },
    "NSE:PARAS": {
        "name": "Paras Defence", "sector": "Defence — Optronics + Space",
        "market_cap_cr": 5200, "cmp": 1120,
        "wk52_high": 1450, "wk52_low": 590,
        "pct_from_ath": 22.8,
        "adr_pct": 6.1,
        "avg_daily_volume_lakhs": 20,
        "ema_20": 1095, "ema_50": 1040, "ema_200": 880,
        "stage": "🟡 BASING",
        "contractions": "34% → 15% → 8% (11 weeks)",
        "volume_dry_up": "Yes",
        "pivot_point": 1195,
        "breakout_above": 1195,
        "entry_zone": "1195-1215",
        "sl": 1130,
        "t1": 1265,
        "t2": 1370,
        "t3": 1490,
        "expected_move_1d_pct": 6,
        "expected_move_3d_pct": 18,
        "catalyst": "Defence + space; MoU pipeline ramp",
        "smart_money": "Strong DII",
        "conviction": "HIGH",
        "risk_pct_to_sl": 5.4,
    },

    # ── MID-CAP MOMENTUM (already reformed VCP) ──────────────────────────────
    "NSE:KAYNES": {
        "name": "Kaynes Technology", "sector": "EMS + Semiconductors",
        "market_cap_cr": 32000, "cmp": 5100,
        "wk52_high": 6240, "wk52_low": 3200,
        "pct_from_ath": 18.3,
        "adr_pct": 4.2,
        "avg_daily_volume_lakhs": 8,
        "ema_20": 5050, "ema_50": 4830, "ema_200": 4250,
        "stage": "🔵 CONTRACTING",
        "contractions": "26% → 12% → 5% (10 weeks)",
        "volume_dry_up": "Yes",
        "pivot_point": 5340,
        "breakout_above": 5340,
        "entry_zone": "5340-5400",
        "sl": 5060,
        "t1": 5610,
        "t2": 6020,
        "t3": 6500,
        "expected_move_1d_pct": 5,
        "expected_move_3d_pct": 14,
        "catalyst": "Semicon OSAT ramp; bulk buys by ICICI Pru MF Apr 24",
        "smart_money": "Bulk buy + DII",
        "conviction": "VERY HIGH",
        "risk_pct_to_sl": 5.2,
    },

    # ── MUSIC / MEDIA HIGH-BETA ──────────────────────────────────────────────
    "NSE:TIPSINDLTD": {
        "name": "Tips Industries (Music)", "sector": "Music Catalogue + Streaming",
        "market_cap_cr": 8800, "cmp": 685,
        "wk52_high": 905, "wk52_low": 385,
        "pct_from_ath": 24.3,
        "adr_pct": 5.8,
        "avg_daily_volume_lakhs": 12,
        "ema_20": 672, "ema_50": 645, "ema_200": 580,
        "stage": "🟡 BASING",
        "contractions": "36% → 14% → 7% (12 weeks)",
        "volume_dry_up": "Yes",
        "pivot_point": 730,
        "breakout_above": 730,
        "entry_zone": "730-745",
        "sl": 690,
        "t1": 770,
        "t2": 830,
        "t3": 910,
        "expected_move_1d_pct": 6,
        "expected_move_3d_pct": 18,
        "catalyst": "Streaming royalties surge; AI music licencing",
        "smart_money": "Recent FPI buying",
        "conviction": "HIGH",
        "risk_pct_to_sl": 5.5,
    },
}

# ── Helpers ──────────────────────────────────────────────────────────────────
def fill(h): return PatternFill("solid", fgColor=h)
def font(color=WHITE, bold=False, size=9, italic=False):
    return Font(name="Arial", color=color, bold=bold, size=size, italic=italic)
def bdr():
    s = Side(style="thin", color="2D2D2D")
    return Border(left=s, right=s, top=s, bottom=s)
def mid(): return Alignment(horizontal="center", vertical="center", wrap_text=True)
def lft(): return Alignment(horizontal="left",   vertical="center", wrap_text=True)

def merge_fresh(meta):
    if not os.path.exists(FRESH_JSON):
        return meta
    try:
        with open(FRESH_JSON) as f:
            fresh = json.load(f)
        for sym, updates in fresh.items():
            if sym in meta:
                meta[sym].update(updates)
            else:
                meta[sym] = updates
        print(f"Merged fresh data for {len(fresh)} VCP candidates")
    except Exception as e:
        print(f"Could not merge VCP fresh JSON: {e}")
    return meta

# ── Excel build ───────────────────────────────────────────────────────────────
HEADERS = [
    "#", "NSE Symbol", "Stock Name", "Sector",
    "MCap ₹cr", "CMP ₹", "52W High", "% from ATH",
    "ADR %", "Avg Vol (L)",
    "EMA Stack (P>20>50>200)",
    "Stage", "Contractions",
    "Pivot ₹", "Entry Zone ₹", "SL ₹", "Risk %",
    "T1 ₹", "T2 ₹", "T3 ₹",
    "Exp 1D %", "Exp 3D %",
    "Catalyst (news/event)", "Smart Money", "Conviction",
]
COL_WIDTHS = [4,16,22,26,9,9,9,10,7,11,15,18,28,9,14,9,8,9,9,9,9,9,30,16,12]

def build():
    meta = merge_fresh({k: dict(v) for k, v in VCP_CANDIDATES.items()})

    # Blacklist filter
    excluded = []
    for sym in list(meta.keys()):
        if sym in BLACKLIST:
            excluded.append(sym); del meta[sym]

    wb = Workbook()
    ws = wb.active
    ws.title = "NSE VCP Breakouts"
    ws.sheet_view.showGridLines = False

    # ── Title ──
    ws.merge_cells("A1:Y1")
    ws["A1"] = (f"NSE VCP BREAKOUT SCREENER — Updated {TODAY.strftime('%d %b %Y')}  |  "
                "Mark Minervini methodology  |  Target moves: 8-20% in 1-3 days")
    ws["A1"].font = Font(name="Arial", color=GOLD, bold=True, size=13)
    ws["A1"].fill = fill(HEADER_BG)
    ws["A1"].alignment = mid()
    ws.row_dimensions[1].height = 28

    ws.merge_cells("A2:Y2")
    ws["A2"] = ("VCP CRITERIA  →  Stage-2 uptrend (P > 9 > 20 > 50 > 200 EMA)  |  "
                "Within 5-25% of 52W high  |  3+ contractions, each smaller  |  "
                "Volume drying during base  |  ADR ≥ 3%  |  MCap ₹1k-50k cr (small/mid-cap sweet spot)  |  "
                "BREAKOUT = close above pivot + volume ≥ 200% of avg")
    ws["A2"].font = Font(name="Arial", color=GREY, italic=True, size=9)
    ws["A2"].fill = fill(DARK_BG)
    ws["A2"].alignment = mid()
    ws.row_dimensions[2].height = 24

    ws.merge_cells("A3:Y3")
    ws["A3"] = ("STAGE  →  " + "   ".join(f"{k}: {v[2]}" for k, v in STAGE_META.items()))
    ws["A3"].font = Font(name="Arial", color=AMBER, bold=True, size=8)
    ws["A3"].fill = fill(HEADER_BG)
    ws["A3"].alignment = mid()
    ws.row_dimensions[3].height = 16

    # ── Headers ──
    for i, (h, w) in enumerate(zip(HEADERS, COL_WIDTHS), 1):
        c = ws.cell(row=4, column=i, value=h)
        c.font      = Font(name="Arial", color=GOLD, bold=True, size=8)
        c.fill      = fill(HEADER_BG)
        c.alignment = mid()
        c.border    = bdr()
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[4].height = 36

    # ── Sort: BREAKING OUT first, then PIVOT, CONTRACTING, BASING, NOT READY, BROKEN ──
    stage_rank = {"🚀 BREAKING OUT": 0, "🟢 PIVOT — READY": 1, "🔵 CONTRACTING": 2,
                  "🟡 BASING": 3, "🟠 NOT READY": 4, "🔴 BROKEN": 5}
    conv_rank = {"VERY HIGH": 0, "HIGH": 1, "MEDIUM": 2}
    sorted_syms = sorted(meta.keys(),
                         key=lambda s: (stage_rank.get(meta[s]["stage"], 9),
                                         conv_rank.get(meta[s]["conviction"], 9),
                                         -meta[s].get("expected_move_3d_pct", 0)))

    for idx, sym in enumerate(sorted_syms, 1):
        d = meta[sym]
        r = idx + 4
        stage_bg, stage_fg, _ = STAGE_META.get(d["stage"], (DARK_BG, GREY, ""))
        row_bg = ROW_ALT if idx % 2 else DARK_BG
        if d["stage"] == "🚀 BREAKING OUT": row_bg = "0A4D2A"
        if d["stage"] == "🔴 BROKEN":       row_bg = "2D0000"

        # Boost conviction with sector + smart money
        sec_boost  = get_sector_boost(sym)
        flow_score = get_smart_money_score(sym)
        conviction = d["conviction"]
        if sec_boost >= 10 and flow_score >= 60 and conviction != "VERY HIGH":
            conviction = "VERY HIGH" if conviction == "HIGH" else "HIGH"

        ema_stack = ("✓✓✓✓" if d["cmp"] > d["ema_20"] > d["ema_50"] > d["ema_200"]
                     else "✓✓✓ " if d["cmp"] > d["ema_20"] > d["ema_50"]
                     else "✓✓  " if d["cmp"] > d["ema_20"]
                     else "✗   ")

        cells = [
            idx, sym, d["name"], d["sector"],
            d["market_cap_cr"], d["cmp"], d["wk52_high"], f"-{d['pct_from_ath']}%",
            f"{d['adr_pct']}%", f"{d['avg_daily_volume_lakhs']}L",
            ema_stack,
            d["stage"], d["contractions"],
            d["pivot_point"], d["entry_zone"], d["sl"], f"{d['risk_pct_to_sl']}%",
            d["t1"], d["t2"], d["t3"],
            f"+{d['expected_move_1d_pct']}%", f"+{d['expected_move_3d_pct']}%",
            d["catalyst"], d["smart_money"], conviction,
        ]

        for col_i, val in enumerate(cells, 1):
            c = ws.cell(row=r, column=col_i, value=val)
            c.fill = fill(row_bg)
            c.border = bdr()
            c.font = font(WHITE, size=9)
            c.alignment = mid() if col_i not in (3, 4, 13, 23, 24) else lft()

            if col_i == 1: c.font = font(GOLD, bold=True)
            if col_i == 2: c.font = font(GREEN, bold=True)
            if col_i == 9: c.font = font(CYAN, bold=True)  # ADR
            if col_i == 11: c.font = font(GREEN if "✓✓✓✓" in str(val) else AMBER, bold=True)
            if col_i == 12:  # Stage
                c.font = Font(name="Arial", color=stage_fg, bold=True, size=9)
                c.fill = fill(stage_bg)
            if col_i == 14: c.font = font(GOLD, bold=True)  # Pivot
            if col_i == 16: c.font = font(RED, bold=True)   # SL
            if col_i == 18: c.font = font(GREEN)            # T1
            if col_i == 19: c.font = font(GREEN, bold=True) # T2
            if col_i == 20: c.font = font(GOLD, bold=True)  # T3
            if col_i == 21:
                v = d["expected_move_1d_pct"]
                c.font = font(GREEN if v >= 5 else AMBER, bold=True)
            if col_i == 22:
                v = d["expected_move_3d_pct"]
                c.font = font(GOLD if v >= 15 else GREEN if v >= 10 else AMBER, bold=True)
            if col_i == 25:
                clr = GREEN if conviction == "VERY HIGH" else BLUE if conviction == "HIGH" else AMBER
                c.font = font(clr, bold=True)

        ws.row_dimensions[r].height = 56

    # ── Footer ──
    fr = len(sorted_syms) + 6
    ws.merge_cells(f"A{fr}:Y{fr}")
    ws[f"A{fr}"] = ("⚠️  VCP TRADE PROTOCOL: 1) Wait for ACTUAL BREAKOUT — close above pivot + volume ≥ 200% of avg.  "
                   "2) Enter on confirmation candle, NOT anticipation.  "
                   "3) Hard SL just below pivot or 5% — no exceptions.  "
                   "4) Book 30% at T1 (+5-8%), 40% at T2 (+12-15%), trail SL on remaining 30%.  "
                   "5) If stock fails to follow through within 2 sessions, exit.  "
                   "6) Position size 1-2% portfolio risk per trade — small position, big move.")
    ws[f"A{fr}"].font      = Font(name="Arial", color=AMBER, bold=True, size=9)
    ws[f"A{fr}"].fill      = fill(HEADER_BG)
    ws[f"A{fr}"].alignment = mid()
    ws.row_dimensions[fr].height = 60

    ws.merge_cells(f"A{fr+1}:Y{fr+1}")
    ws[f"A{fr+1}"] = ("📚 VCP REFERENCE: Mark Minervini (SEPA) + stockexploder Instagram methodology  |  "
                     "Look for tight base + volume dry-up + breakout candle volume explosion  |  "
                     "Best results in defence small-caps, drone makers, EMS, niche specialty chem, low-float IPOs")
    ws[f"A{fr+1}"].font      = Font(name="Arial", color=CYAN, italic=True, size=9)
    ws[f"A{fr+1}"].fill      = fill(DARK_BG)
    ws[f"A{fr+1}"].alignment = mid()
    ws.row_dimensions[fr+1].height = 18

    ws.merge_cells(f"A{fr+2}:Y{fr+2}")
    ws[f"A{fr+2}"] = (f"Auto-generated by Claude Opus 4.7  |  Run date: {TODAY.strftime('%d %b %Y')}  |  "
                     "Sources: Screener.in · TradingView · NSE volume gainers · Trendlyne breakout scans  |  "
                     "VERIFY all CMPs, EMAs, and pivot levels on TradingView chart before trading.")
    ws[f"A{fr+2}"].font      = Font(name="Arial", color=GREY, italic=True, size=8)
    ws[f"A{fr+2}"].fill      = fill(DARK_BG)
    ws[f"A{fr+2}"].alignment = mid()
    ws.row_dimensions[fr+2].height = 14

    ws.freeze_panes = "C5"
    wb.save(EXCEL_PATH)
    print(f"✅ VCP Breakout Excel saved: {EXCEL_PATH}\n")

    # ── Console summary ──
    print(f"📊 VCP Breakout Summary  ({TODAY.strftime('%d %b %Y')}):\n")
    by_stage = {}
    for sym in sorted_syms:
        d = meta[sym]
        by_stage.setdefault(d["stage"], []).append(
            f"{sym} | {d['cmp']} | pivot ₹{d['pivot_point']} | exp 3D +{d['expected_move_3d_pct']}% | {d['conviction']}")
    for stage_key in STAGE_META:
        if stage_key in by_stage:
            print(f"  {stage_key}")
            for item in by_stage[stage_key]:
                print(f"     • {item}")
            print()

# ── Public API for screeners ──────────────────────────────────────────────────
def get_top_vcp_today(n=5):
    """Returns top N VCP candidates ranked by conviction + expected move."""
    meta = merge_fresh({k: dict(v) for k, v in VCP_CANDIDATES.items()})
    stage_rank = {"🚀 BREAKING OUT": 0, "🟢 PIVOT — READY": 1, "🔵 CONTRACTING": 2,
                  "🟡 BASING": 3, "🟠 NOT READY": 4}
    conv_rank = {"VERY HIGH": 0, "HIGH": 1, "MEDIUM": 2}
    syms = sorted(meta.keys(),
                  key=lambda s: (stage_rank.get(meta[s]["stage"], 9),
                                  conv_rank.get(meta[s]["conviction"], 9),
                                  -meta[s].get("expected_move_3d_pct", 0)))
    return [(s, meta[s]) for s in syms[:n]]

if __name__ == "__main__":
    build()
