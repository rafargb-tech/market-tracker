"""
Market Tracker v2 - Índices, Bonos, Commodities + Macro Dashboard
Requiere: pip install yfinance openpyxl pandas_datareader
"""

import yfinance as yf
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from datetime import datetime, timedelta

# ── UNIVERSO DE ACTIVOS ───────────────────────────────────────────────────────

SECTIONS = [
    ("MAJOR INDICES", [
        ("S&P 500",               "^GSPC"),
        ("MSCI World",            "URTH"),
        ("NASDAQ Composite",      "^IXIC"),
        ("Euro STOXX 50",         "^STOXX50E"),
        ("MSCI Emerging Markets", "EEM"),
        ("Russell 2000",          "^RUT"),
        ("S&P 500 Equal Weight",  "RSP"),
    ]),
    ("MAG7", [
        ("NVIDIA",         "NVDA"),
        ("Apple",          "AAPL"),
        ("Alphabet",       "GOOGL"),
        ("Microsoft",      "MSFT"),
        ("Amazon",         "AMZN"),
        ("Meta Platforms", "META"),
        ("Tesla",          "TSLA"),
    ]),
    ("US TREASURIES", [
        ("US 2Y Yield",    "^IRX"),
        ("US 5Y Yield",    "^FVX"),
        ("US 10Y Yield",   "^TNX"),
        ("US 30Y Yield",   "^TYX"),
        ("TLT (20Y+ ETF)", "TLT"),
        ("IEF (7-10Y ETF)","IEF"),
        ("SHY (1-3Y ETF)", "SHY"),
    ]),
    ("GLOBAL BONDS", [
        ("Bund 10Y (Germany)", "^DE10YB"),
        ("Gilt 10Y (UK)",      "^GB10YB"),
        ("JGB 10Y (Japan)",    "^JP10YB"),
        ("BTP 10Y (Italy)",    "^IT10YB"),
        ("OAT 10Y (France)",   "^FR10YB"),
        ("Bonos 10Y (Spain)",  "^ES10YB"),
    ]),
    ("COMMODITIES", [
        ("Gold",        "GC=F"),
        ("Silver",      "SI=F"),
        ("Oil (WTI)",   "CL=F"),
        ("Oil (Brent)", "BZ=F"),
        ("Natural Gas", "NG=F"),
        ("Copper",      "HG=F"),
        ("Bitcoin",     "BTC-USD"),
        ("Ethereum",    "ETH-USD"),
    ]),
    ("EUROPE", [
        ("United Kingdom (FTSE 100)", "^FTSE"),
        ("France (CAC 40)",           "^FCHI"),
        ("Germany (DAX)",             "^GDAXI"),
        ("Netherlands (AEX)",         "^AEX"),
        ("Spain (IBEX 35)",           "^IBEX"),
    ]),
    ("ASIA", [
        ("Japan",      "^N225"),
        ("South Korea","^KS11"),
        ("India",      "^BSESN"),
        ("China",      "MCHI"),
        ("Hong Kong",  "^HSI"),
    ]),
    ("LATAM", [
        ("Brazil",    "EWZ"),
        ("Mexico",    "EWW"),
        ("Argentina", "ARGT"),
    ]),
    ("US SECTORS", [
        ("Technology",             "XLK"),
        ("Healthcare",             "XLV"),
        ("Financials",             "XLF"),
        ("Consumer Discretionary", "XLY"),
        ("Communication Services", "XLC"),
        ("Industrials",            "XLI"),
        ("Consumer Staples",       "XLP"),
        ("Energy",                 "XLE"),
        ("Utilities",              "XLU"),
        ("Real Estate",            "XLRE"),
        ("Materials",              "XLB"),
    ]),
    ("EU SECTORS", [
        ("EU Banks",                  "EXV1.DE"),
        ("EU Healthcare",             "EXH1.DE"),
        ("EU Industrials",            "EXH2.DE"),
        ("EU Energy",                 "EXH8.DE"),
        ("EU Technology",             "EXV4.DE"),
        ("EU Consumer Discretionary", "EXH3.DE"),
        ("EU Consumer Staples",       "EXH4.DE"),
        ("EU Telecom",                "EXH6.DE"),
        ("EU Utilities",              "EXH7.DE"),
        ("EU Materials",              "EXH5.DE"),
        ("EU Real Estate",            "IPRP.AS"),
    ]),
]

# ── MACRO ─────────────────────────────────────────────────────────────────────

MACRO_DATA = {
    "UNITED STATES": [
        ("Fed Funds Rate",    "FEDFUNDS",              "%",   "FRED", "Target upper bound"),
        ("CPI YoY",           "CPIAUCSL",              "%",   "FRED", "All items, not seas. adj."),
        ("Core CPI YoY",      "CPILFESL",              "%",   "FRED", "Ex food & energy"),
        ("GDP Growth QoQ",    "A191RL1Q225SBEA",       "%",   "FRED", "Real GDP, annualised"),
        ("Unemployment Rate", "UNRATE",                "%",   "FRED", "U-3 rate"),
        ("10Y-2Y Spread",     "T10Y2Y",                "pts", "FRED", "Negative = inverted curve"),
        ("VIX",               "^VIX",                  "pts", "YF",   ">30 = high fear"),
        ("DXY (USD Index)",   "DX-Y.NYB",              "pts", "YF",   "Trade-weighted USD"),
    ],
    "EURO ZONE": [
        ("ECB Deposit Rate",  "ECBDFR",                "%",   "FRED", "ECB deposit facility rate"),
        ("CPI YoY (EA)",      "CP0000EZ19M086NEST",    "%",   "FRED", "Euro Area HICP"),
        ("Unemployment (EA)", "LRHUTTTTEZM156S",       "%",   "FRED", "Euro Area unemployment"),
        ("GDP Growth QoQ",    "CLVMNACSCAB1GQEA19",    "%",   "FRED", "Euro Area real GDP"),
        ("EUR/USD",           "EURUSD=X",              "fx",  "YF",   "Spot rate"),
        ("GBP/USD",           "GBPUSD=X",              "fx",  "YF",   "Spot rate"),
        ("USD/JPY",           "JPY=X",                 "fx",  "YF",   "JPY per USD"),
    ],
}

# ── CÁLCULOS ──────────────────────────────────────────────────────────────────

def get_prices(tickers):
    end   = datetime.today()
    start = end - timedelta(days=365 * 5 + 10)
    raw   = yf.download(tickers, start=start, end=end, auto_adjust=True, progress=False)
    return raw["Close"]

def calc_return(series, days):
    s = series.dropna()
    if len(s) < 2: return None
    past = s.iloc[-days] if len(s) > days else s.iloc[0]
    return (s.iloc[-1] - past) / past if past != 0 else None

def ytd_return(series):
    s = series.dropna()
    if len(s) < 2: return None
    ytd_s = s[s.index >= datetime(datetime.today().year, 1, 1)]
    if len(ytd_s) < 1: return None
    return (s.iloc[-1] - ytd_s.iloc[0]) / ytd_s.iloc[0] if ytd_s.iloc[0] != 0 else None

def pct_from_low_high(series):
    s = series.dropna().iloc[-252:]
    if len(s) == 0: return None, None
    p = series.dropna().iloc[-1]
    lo, hi = s.min(), s.max()
    return ((p-lo)/lo if lo!=0 else None, (p-hi)/hi if hi!=0 else None)

def build_market_rows(prices_df):
    rows = []
    for section_name, assets in SECTIONS:
        rows.append(("HEADER", section_name))
        for name, ticker in assets:
            if ticker not in prices_df.columns:
                rows.append((name, ticker, None, None, None, None, None, None, None, None, None))
                continue
            s = prices_df[ticker]
            price = float(s.dropna().iloc[-1]) if len(s.dropna()) > 0 else None
            pct_low, pct_high = pct_from_low_high(s)
            rows.append((name, ticker, price, pct_low, pct_high,
                         calc_return(s,1), calc_return(s,5), calc_return(s,21),
                         ytd_return(s), calc_return(s,252), calc_return(s,252*5)))
    return rows

def get_fred_series(series_id):
    try:
        import pandas_datareader.data as web
        df = web.DataReader(series_id, "fred",
                            start=datetime.today()-timedelta(days=800),
                            end=datetime.today())
        s = df[series_id].dropna()
        if len(s) == 0: return None, None, None
        latest   = float(s.iloc[-1])
        chg_last = float(s.iloc[-1] - s.iloc[-2]) if len(s) > 1 else None
        chg_yoy  = float(s.iloc[-1] - s.iloc[-13]) if len(s) > 13 else None
        return latest, chg_last, chg_yoy
    except Exception:
        return None, None, None

def get_yf_macro(ticker):
    try:
        df = yf.download(ticker, period="2y", auto_adjust=True, progress=False)
        if df.empty: return None, None, None
        s = df["Close"].dropna()
        latest   = float(s.iloc[-1])
        chg_last = float(s.iloc[-1] - s.iloc[-2]) if len(s) > 1 else None
        chg_yoy  = float(s.iloc[-1] - s.iloc[-252]) if len(s) > 252 else None
        return latest, chg_last, chg_yoy
    except Exception:
        return None, None, None

def build_macro_data():
    results = {}
    for region, indicators in MACRO_DATA.items():
        results[region] = []
        for name, series_id, unit, source, note in indicators:
            print(f"  📊 {name}...")
            if source == "YF":
                val, chg_last, chg_yoy = get_yf_macro(series_id)
            else:
                val, chg_last, chg_yoy = get_fred_series(series_id)
            results[region].append((name, unit, val, chg_last, chg_yoy, note))
    return results

# ── ESTILOS ───────────────────────────────────────────────────────────────────

DARK_BG    = "1C2634"
HEADER_BG  = "2C3E50"
SECTION_BG = "243447"
ACCENT     = "1A6B9A"
WHITE      = "FFFFFF"
LIGHT_GRAY = "D6DCE4"
MID_GRAY   = "8496A9"

def fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def fnt(bold=False, color=WHITE, size=9, italic=False):
    return Font(bold=bold, color=color, size=size, name="Arial", italic=italic)

def center():
    return Alignment(horizontal="center", vertical="center")

def left(indent=0):
    return Alignment(horizontal="left", vertical="center", indent=indent)

def color_ret(val):
    if val is None:        return HEADER_BG
    if val >=  0.10:       return "1A7A3A"
    if val >=  0.05:       return "27AE60"
    if val >=  0.02:       return "52BE80"
    if val >=  0.00:       return "A9DFBF"
    if val >= -0.02:       return "F1948A"
    if val >= -0.05:       return "E74C3C"
    if val >= -0.10:       return "C0392B"
    return "922B21"

def text_ret(val):
    if val is None: return MID_GRAY
    return "2C2C2C" if abs(val) < 0.02 else WHITE

def color_chg(val):
    if val is None or val == 0: return SECTION_BG
    return "27AE60" if val > 0 else "E74C3C"

def fmt_pct(val):
    if val is None: return "–"
    return f"{'+'if val>0 else ''}{val*100:.1f}%"

def fmt_price(val):
    if val is None: return "–"
    if val >= 10000: return f"{val:,.0f}"
    if val >= 100:   return f"{val:,.1f}"
    if val >= 1:     return f"{val:,.2f}"
    return f"{val:.4f}"

def fmt_macro(val, unit):
    if val is None: return "N/A"
    if unit == "%":   return f"{val:.2f}%"
    if unit == "pts": return f"{val:.2f}"
    if unit == "fx":  return f"{val:.4f}"
    return f"{val:.2f}"

def fmt_chg(val, unit):
    if val is None: return "–"
    sign = "+" if val > 0 else ""
    if unit == "%":   return f"{sign}{val:.2f}pp"
    if unit == "pts": return f"{sign}{val:.2f}"
    if unit == "fx":  return f"{sign}{val:.4f}"
    return f"{sign}{val:.2f}"

# ── SHEET 1: MARKETS ──────────────────────────────────────────────────────────

MKT_HEADERS = ["Name", "Ticker", "Price", "vs 52W Low", "vs 52W High",
               "1D", "1W", "1M", "YTD", "1Y", "5Y"]
MKT_WIDTHS  = [28, 9, 11, 11, 11, 8, 8, 8, 8, 8, 8]

def write_market_sheet(ws, rows, today_str):
    ws.sheet_view.showGridLines = False

    ws.merge_cells("A1:H1")
    ws["A1"] = "Market Tracker"
    ws["A1"].font = fnt(bold=True, size=14); ws["A1"].fill = fill(DARK_BG); ws["A1"].alignment = left(1)

    ws.merge_cells("I1:K1")
    ws["I1"] = today_str
    ws["I1"].font = fnt(bold=True, color=LIGHT_GRAY, size=11); ws["I1"].fill = fill(DARK_BG); ws["I1"].alignment = center()

    ws.merge_cells("A2:K2")
    ws["A2"] = "All returns in local currency  |  Bond tickers show yield level or ETF price"
    ws["A2"].font = fnt(italic=True, color=MID_GRAY, size=8); ws["A2"].fill = fill(DARK_BG); ws["A2"].alignment = left(1)

    for rng, label in [("A3:B3","Overview"),("C3:E3","52W Range"),("F3:K3","Returns")]:
        ws.merge_cells(rng)
        c = ws[rng.split(":")[0]]
        c.value = label; c.font = fnt(bold=True, color=LIGHT_GRAY, size=8)
        c.fill = fill(HEADER_BG); c.alignment = center()

    for ci, (h, w) in enumerate(zip(MKT_HEADERS, MKT_WIDTHS), 1):
        c = ws.cell(row=4, column=ci, value=h)
        c.font = fnt(bold=True, color=LIGHT_GRAY, size=8); c.fill = fill(HEADER_BG); c.alignment = center()
        ws.column_dimensions[get_column_letter(ci)].width = w

    ws.row_dimensions[1].height = 22
    ws.row_dimensions[3].height = 13
    ws.row_dimensions[4].height = 15

    er = 5
    for rd in rows:
        if rd[0] == "HEADER":
            ws.merge_cells(f"A{er}:K{er}")
            c = ws.cell(row=er, column=1, value=rd[1])
            c.font = fnt(bold=True, color=LIGHT_GRAY, size=8); c.fill = fill(SECTION_BG); c.alignment = left(1)
            ws.row_dimensions[er].height = 13; er += 1; continue

        name, ticker, price, pct_low, pct_high, r1d, r1w, r1m, rytd, r1y, r5y = rd
        vals = [name, ticker, fmt_price(price), fmt_pct(pct_low), fmt_pct(pct_high),
                fmt_pct(r1d), fmt_pct(r1w), fmt_pct(r1m), fmt_pct(rytd), fmt_pct(r1y), fmt_pct(r5y)]
        rets = [None,None,None,None,None, r1d, r1w, r1m, rytd, r1y, r5y]

        for ci, (v, r) in enumerate(zip(vals, rets), 1):
            c = ws.cell(row=er, column=ci, value=v)
            c.alignment = left(1) if ci <= 2 else center()
            if   ci == 1: c.fill=fill(DARK_BG);    c.font=fnt(color=LIGHT_GRAY, size=8)
            elif ci == 2: c.fill=fill(DARK_BG);    c.font=fnt(color=MID_GRAY, size=8, italic=True)
            elif ci == 3: c.fill=fill(DARK_BG);    c.font=fnt(size=8)
            elif ci in (4,5): c.fill=fill(SECTION_BG); c.font=fnt(color=LIGHT_GRAY, size=8)
            else:
                c.fill = fill(color_ret(r))
                c.font = fnt(color=text_ret(r), size=8, bold=(r is not None and abs(r)>0.05))
        ws.row_dimensions[er].height = 14; er += 1

    ws.freeze_panes = "A5"

# ── SHEET 2: MACRO ────────────────────────────────────────────────────────────

MAC_HEADERS = ["Indicator", "Unit", "Latest", "Chg (last period)", "Chg (YoY/12m)", "Note"]
MAC_WIDTHS  = [30, 7, 13, 18, 15, 38]

def write_macro_sheet(ws, macro_data, today_str):
    ws.sheet_view.showGridLines = False

    ws.merge_cells("A1:E1")
    ws["A1"] = "Macro Dashboard"
    ws["A1"].font = fnt(bold=True, size=14); ws["A1"].fill = fill(DARK_BG); ws["A1"].alignment = left(1)

    ws["F1"] = today_str
    ws["F1"].font = fnt(bold=True, color=LIGHT_GRAY, size=11); ws["F1"].fill = fill(DARK_BG); ws["F1"].alignment = center()

    ws.merge_cells("A2:F2")
    ws["A2"] = "FRED series via pandas_datareader  |  YF = Yahoo Finance  |  Data may lag 1-2 months for official stats"
    ws["A2"].font = fnt(italic=True, color=MID_GRAY, size=8); ws["A2"].fill = fill(DARK_BG); ws["A2"].alignment = left(1)

    for ci, (h, w) in enumerate(zip(MAC_HEADERS, MAC_WIDTHS), 1):
        ws.column_dimensions[get_column_letter(ci)].width = w

    ws.row_dimensions[1].height = 22
    ws.row_dimensions[2].height = 13

    er = 3
    for region, indicators in macro_data.items():
        ws.merge_cells(f"A{er}:F{er}")
        c = ws.cell(row=er, column=1, value=region)
        c.font = fnt(bold=True, color=LIGHT_GRAY, size=9); c.fill = fill(SECTION_BG); c.alignment = left(1)
        ws.row_dimensions[er].height = 15; er += 1

        for ci, h in enumerate(MAC_HEADERS, 1):
            c = ws.cell(row=er, column=ci, value=h)
            c.font = fnt(bold=True, color=LIGHT_GRAY, size=8); c.fill = fill(HEADER_BG); c.alignment = center()
        ws.row_dimensions[er].height = 14; er += 1

        for name, unit, val, chg_last, chg_yoy, note in indicators:
            row_vals = [name, unit, fmt_macro(val, unit), fmt_chg(chg_last, unit), fmt_chg(chg_yoy, unit), note]
            for ci, rv in enumerate(row_vals, 1):
                c = ws.cell(row=er, column=ci, value=rv)
                c.alignment = left(1) if ci in (1, 6) else center()
                if   ci == 1: c.fill=fill(DARK_BG);    c.font=fnt(color=LIGHT_GRAY, size=8)
                elif ci == 2: c.fill=fill(DARK_BG);    c.font=fnt(color=MID_GRAY, size=8)
                elif ci == 3: c.fill=fill(DARK_BG);    c.font=fnt(bold=True, size=9)
                elif ci == 4:
                    bg = color_chg(chg_last)
                    c.fill=fill(bg); c.font=fnt(color=WHITE if bg!=SECTION_BG else MID_GRAY, size=8)
                elif ci == 5:
                    bg = color_chg(chg_yoy)
                    c.fill=fill(bg); c.font=fnt(color=WHITE if bg!=SECTION_BG else MID_GRAY, size=8)
                else: c.fill=fill(DARK_BG); c.font=fnt(color=MID_GRAY, size=8, italic=True)
            ws.row_dimensions[er].height = 14; er += 1

        er += 1

    ws.freeze_panes = "A3"

# ── MAIN ──────────────────────────────────────────────────────────────────────

def main():
    today_str = datetime.today().strftime("%d-%b-%y")
    output    = f"market_tracker_{datetime.today().strftime('%Y%m%d')}.xlsx"

    print("📡 Descargando precios de mercado...")
    all_tickers = [t for _, assets in SECTIONS for _, t in assets]
    prices = get_prices(all_tickers)

    print("🔢 Calculando retornos...")
    market_rows = build_market_rows(prices)

    print("🌍 Descargando datos macro...")
    macro_data = build_macro_data()

    print("📝 Generando Excel...")
    wb  = openpyxl.Workbook()
    ws1 = wb.active
    ws1.title = "Markets"
    write_market_sheet(ws1, market_rows, today_str)

    ws2 = wb.create_sheet("Macro")
    write_macro_sheet(ws2, macro_data, today_str)

    wb.save(output)
    print(f"\n✅ Guardado: {output}")
    print("   Pestaña 1 → Markets  |  Pestaña 2 → Macro")

if __name__ == "__main__":
    main()
