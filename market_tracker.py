"""
Market Tracker v2 - Índices, Bonos, Commodities + Macro Dashboard
Requiere: pip install yfinance openpyxl
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
    ("US TREASURIES (ETFs)", [
        ("TLT (20Y+ ETF)",  "TLT"),
        ("IEF (7-10Y ETF)", "IEF"),
        ("IEI (3-7Y ETF)",  "IEI"),
        ("SHY (1-3Y ETF)",  "SHY"),
        ("BIL (0-3M ETF)",  "BIL"),
    ]),
    ("GLOBAL BONDS (ETFs)", [
        ("Bund 10Y · iShares (EUR)", "EXX6.DE"),   # iShares Bund 7-10Y
        ("Gilt 10Y · iShares (GBP)", "IGLT.L"),    # iShares Core UK Gilts
        ("JGB · iShares (JPY)",      "2561.T"),     # iShares JPY Govt Bond
        ("Italy BTP · iShares",      "ITPS.MI"),   # iShares BTP
        ("EU Govt Bond · iShares",   "IEAG.AS"),   # iShares Core Euro Govt
        ("EM Bond · iShares (USD)",  "EMB"),        # iShares EM USD Bond
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
    "US YIELD CURVE": [
        ("US 3M Yield",  "DGS3MO", "%", "FRED", "T-Bill 3 meses"),
        ("US 2Y Yield",  "DGS2",   "%", "FRED", "Treasury 2 años"),
        ("US 5Y Yield",  "DGS5",   "%", "FRED", "Treasury 5 años"),
        ("US 10Y Yield", "DGS10",  "%", "FRED", "Treasury 10 años"),
        ("US 30Y Yield", "DGS30",  "%", "FRED", "Treasury 30 años"),
        ("10Y - 2Y",     "T10Y2Y", "pts", "FRED", "Spread curva. Negativo = invertida"),
        ("10Y - 3M",     "T10Y3M", "pts", "FRED", "Spread curva corta"),
    ],
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
    """Descarga datos de FRED via API REST pública (sin API key ni pandas_datareader)."""
    try:
        import urllib.request, json
        start = (datetime.today() - timedelta(days=800)).strftime("%Y-%m-%d")
        url = (f"https://fred.stlouisfed.org/graph/fredgraph.csv?id={series_id}"
               f"&vintage_date={datetime.today().strftime('%Y-%m-%d')}")
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=15) as r:
            lines = r.read().decode().strip().split("\n")[1:]  # skip header
        vals = []
        for line in lines:
            parts = line.split(",")
            if len(parts) == 2 and parts[1].strip() not in ("", "."):
                try:
                    vals.append(float(parts[1]))
                except ValueError:
                    pass
        if not vals: return None, None, None
        latest   = vals[-1]
        chg_last = vals[-1] - vals[-2]  if len(vals) > 1  else None
        chg_yoy  = vals[-1] - vals[-13] if len(vals) > 13 else None
        return latest, chg_last, chg_yoy
    except Exception:
        return None, None, None

def get_yf_macro(ticker):
    try:
        df = yf.download(ticker, period="2y", auto_adjust=True, progress=False)
        if df.empty: return None, None, None
        s = df["Close"].dropna()
        if hasattr(s, 'columns'):  # MultiIndex — coger primera columna
            s = s.iloc[:, 0]
        latest   = float(s.iloc[-1].item() if hasattr(s.iloc[-1], 'item') else s.iloc[-1])
        chg_last = float((s.iloc[-1] - s.iloc[-2]).item()) if len(s) > 1 else None
        chg_yoy  = float((s.iloc[-1] - s.iloc[-252]).item()) if len(s) > 252 else None
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

# ── SPI: SECTOR PULSE INVESTING ───────────────────────────────────────────────

# Pesos por fase (Expansión Temprana, Expansión Tardía, Recesión Temprana, Recesión Tardía)
SPI_SECTORS = [
    # (Nombre,           Ticker, ExpT,  ExpL,  RecT,  RecL)
    ("Technology",       "XLK",  0.25,  0.08,  0.05,  0.08),
    ("Cons. Discret.",   "XLY",  0.20,  0.08,  0.05,  0.05),
    ("Communication",    "XLC",  0.15,  0.07,  0.05,  0.05),
    ("Financials",       "XLF",  0.08,  0.20,  0.05,  0.05),
    ("Industrials",      "XLI",  0.08,  0.20,  0.05,  0.05),
    ("Materials",        "XLB",  0.05,  0.15,  0.05,  0.05),
    ("Cons. Staples",    "XLP",  0.05,  0.05,  0.20,  0.15),
    ("Healthcare",       "XLV",  0.05,  0.05,  0.25,  0.15),
    ("Utilities",        "XLU",  0.04,  0.04,  0.20,  0.10),
    ("Energy",           "XLE",  0.03,  0.08,  0.05,  0.15),
    ("Real Estate",      "XLRE", 0.02,  0.00,  0.00,  0.12),
]

PHASE_NAMES  = ["Expansión Temprana", "Expansión Tardía", "Recesión Temprana", "Recesión Tardía"]
PHASE_COLORS = ["27AE60", "F39C12", "E74C3C", "8E44AD"]  # verde, naranja, rojo, morado
PHASE_KEYS   = ["exp_early", "exp_late", "rec_early", "rec_late"]

# Indicadores macro usados para determinar la fase
# Cada indicador devuelve su último valor y su variación (chg)
# Lógica de scoring:
#   gdp_growing + unemp_falling + cpi_low  + rates_low/falling  → Expansión Temprana
#   gdp_growing + unemp_stable  + cpi_high + rates_rising        → Expansión Tardía
#   gdp_falling + unemp_rising  + cpi_high + rates_high          → Recesión Temprana
#   gdp_low     + unemp_high    + cpi_low  + rates_falling       → Recesión Tardía

def detect_cycle_phase(macro_results):
    """
    Determina la fase del ciclo usando los datos FRED ya descargados.
    Devuelve (phase_index 0-3, dict con señales individuales).
    """
    signals = {}

    # Extraer valores de las series FRED ya descargadas
    def get_val(region, name):
        for r, indicators in macro_results.items():
            for n, unit, val, chg_last, chg_yoy, note in indicators:
                if n == name:
                    return val, chg_last
        return None, None

    gdp_val,   gdp_chg   = get_val("UNITED STATES", "GDP Growth QoQ")
    unemp_val, unemp_chg = get_val("UNITED STATES", "Unemployment Rate")
    cpi_val,   cpi_chg   = get_val("UNITED STATES", "CPI YoY")
    fed_val,   fed_chg   = get_val("UNITED STATES", "Fed Funds Rate")

    # Señales booleanas
    gdp_growing   = gdp_val  is not None and gdp_val  > 0
    gdp_strong    = gdp_val  is not None and gdp_val  > 2.5
    unemp_falling = unemp_chg is not None and unemp_chg < 0
    unemp_rising  = unemp_chg is not None and unemp_chg > 0
    unemp_high    = unemp_val is not None and unemp_val > 5.0
    cpi_high      = cpi_val  is not None and cpi_val  > 3.0
    cpi_rising    = cpi_chg  is not None and cpi_chg  > 0
    rates_rising  = fed_chg  is not None and fed_chg  > 0
    rates_falling = fed_chg  is not None and fed_chg  < 0
    rates_high    = fed_val  is not None and fed_val  > 4.0

    signals = {
        "GDP":        ("▲ Creciendo"  if gdp_growing  else "▼ Cayendo",   gdp_val,   gdp_chg),
        "Desempleo":  ("▼ Bajando"    if unemp_falling else "▲ Subiendo",  unemp_val, unemp_chg),
        "CPI":        ("▲ Alto"       if cpi_high      else "✓ Moderado",  cpi_val,   cpi_chg),
        "Fed Funds":  ("▲ Subiendo"   if rates_rising  else ("▼ Bajando" if rates_falling else "→ Estable"), fed_val, fed_chg),
    }

    # Scoring de fases
    score = [0, 0, 0, 0]  # [ExpT, ExpL, RecT, RecL]

    if gdp_growing and unemp_falling and not cpi_high and not rates_high:
        score[0] += 3  # Expansión Temprana clara
    if gdp_strong and cpi_high and rates_rising:
        score[1] += 3  # Expansión Tardía clara
    if not gdp_growing and unemp_rising and cpi_high:
        score[2] += 3  # Recesión Temprana clara
    if not gdp_growing and unemp_high and not cpi_rising and (rates_falling or not rates_high):
        score[3] += 3  # Recesión Tardía clara

    # Señales adicionales de refuerzo
    if gdp_growing:   score[0] += 1; score[1] += 1
    if unemp_falling: score[0] += 1
    if unemp_rising:  score[2] += 1; score[3] += 1
    if cpi_high:      score[1] += 1; score[2] += 1
    if rates_rising:  score[1] += 1
    if rates_falling: score[3] += 1; score[0] += 1

    phase_idx = score.index(max(score))
    return phase_idx, signals

def get_ema200_weekly(ticker, prices_daily):
    """
    Calcula EMA200 semanal a partir de precios diarios.
    Resamplea a semanas y calcula EMA de 200 periodos.
    """
    try:
        if ticker not in prices_daily.columns:
            return None, None
        s = prices_daily[ticker].dropna()
        if len(s) < 10:
            return None, None
        weekly = s.resample("W").last().dropna()
        if len(weekly) < 10:
            return None, None
        ema200 = weekly.ewm(span=200, adjust=False).mean()
        price   = float(s.iloc[-1])
        ema_val = float(ema200.iloc[-1])
        above   = price > ema_val
        pct_vs  = (price - ema_val) / ema_val
        return above, pct_vs
    except Exception:
        return None, None

def build_spi_data(prices_daily, macro_results):
    """Construye todos los datos necesarios para la pestaña SPI."""
    phase_idx, signals = detect_cycle_phase(macro_results)

    sector_data = []
    for name, ticker, w_et, w_el, w_rt, w_rl in SPI_SECTORS:
        weights   = [w_et, w_el, w_rt, w_rl]
        rec_weight = weights[phase_idx]

        above_ema, pct_ema = get_ema200_weekly(ticker, prices_daily)

        # Precio actual y retornos
        price, r1m, r3m, ytd, r1y = None, None, None, None, None
        if ticker in prices_daily.columns:
            s     = prices_daily[ticker].dropna()
            price = float(s.iloc[-1]) if len(s) > 0 else None
            r1m   = calc_return(s, 21)
            r3m   = calc_return(s, 63)
            ytd   = ytd_return(s)
            r1y   = calc_return(s, 252)

        # Señal: SOBREPONDERAR si peso alto Y por debajo de EMA200
        alerta = False
        if above_ema is not None and not above_ema and rec_weight >= 0.10:
            alerta = True

        sector_data.append({
            "name":       name,
            "ticker":     ticker,
            "weights":    weights,
            "rec_weight": rec_weight,
            "price":      price,
            "r1m":        r1m,
            "r3m":        r3m,
            "ytd":        ytd,
            "r1y":        r1y,
            "above_ema":  above_ema,
            "pct_ema":    pct_ema,
            "alerta":     alerta,
        })

    return phase_idx, signals, sector_data

# ── SHEET 3: SPI ──────────────────────────────────────────────────────────────

def write_spi_sheet(ws, phase_idx, signals, sector_data, today_str):
    ws.sheet_view.showGridLines = False

    phase_name  = PHASE_NAMES[phase_idx]
    phase_color = PHASE_COLORS[phase_idx]

    # ── Título ──
    ws.merge_cells("A1:H1")
    ws["A1"] = "Sector Pulse Investing (SPI)"
    ws["A1"].font = fnt(bold=True, size=14); ws["A1"].fill = fill(DARK_BG); ws["A1"].alignment = left(1)
    ws.merge_cells("I1:L1")
    ws["I1"] = today_str
    ws["I1"].font = fnt(bold=True, color=LIGHT_GRAY, size=11); ws["I1"].fill = fill(DARK_BG); ws["I1"].alignment = center()

    # ── Fase detectada ──
    ws.merge_cells("A2:L2")
    ws["A2"] = f"  FASE ACTUAL DEL CICLO:  {phase_name.upper()}"
    ws["A2"].font = Font(bold=True, color=WHITE, size=12, name="Arial")
    ws["A2"].fill = fill(phase_color)
    ws["A2"].alignment = left(1)
    ws.row_dimensions[2].height = 22

    # ── Indicadores macro que determinan la fase ──
    ws.merge_cells("A3:L3")
    ws["A3"] = "  Indicadores del ciclo:"
    ws["A3"].font = fnt(bold=True, color=LIGHT_GRAY, size=8)
    ws["A3"].fill = fill(SECTION_BG); ws["A3"].alignment = left(1)
    ws.row_dimensions[3].height = 13

    # Fila de indicadores
    signal_cols = list(signals.items())
    col = 1
    for i, (ind_name, (status, val, chg)) in enumerate(signal_cols):
        # Nombre indicador
        c = ws.cell(row=4, column=col, value=ind_name)
        c.font = fnt(bold=True, color=MID_GRAY, size=8); c.fill = fill(DARK_BG); c.alignment = center()
        ws.merge_cells(start_row=4, start_column=col, end_row=4, end_column=col+2)

        # Valor y estado
        val_str = f"{val:.1f}" if val is not None else "N/A"
        chg_str = f"({'+' if chg and chg>0 else ''}{chg:.2f})" if chg is not None else ""
        c2 = ws.cell(row=5, column=col, value=f"{val_str} {chg_str}")
        c2.font = fnt(bold=True, size=9); c2.fill = fill(DARK_BG); c2.alignment = center()
        ws.merge_cells(start_row=5, start_column=col, end_row=5, end_column=col+2)

        c3 = ws.cell(row=6, column=col, value=status)
        is_positive = "▲" in status if "Desempleo" not in ind_name else "▼" in status
        c3.font = fnt(bold=True, color="27AE60" if is_positive else "E74C3C", size=8)
        c3.fill = fill(DARK_BG); c3.alignment = center()
        ws.merge_cells(start_row=6, start_column=col, end_row=6, end_column=col+2)

        col += 3

    ws.row_dimensions[4].height = 13
    ws.row_dimensions[5].height = 16
    ws.row_dimensions[6].height = 13

    # ── Espacio ──
    ws.row_dimensions[7].height = 6
    for c in range(1, 13):
        ws.cell(row=7, column=c).fill = fill(DARK_BG)

    # ── Cabeceras tabla de sectores ──
    headers = ["Sector", "Ticker", "Precio",
               "Exp. Temp.", "Exp. Tard.", "Rec. Temp.", "Rec. Tard.",
               "Peso Actual", "1M", "3M", "YTD", "vs EMA200W"]
    widths  = [18, 7, 9, 10, 10, 10, 10, 11, 8, 8, 8, 12]

    for ci, (h, w) in enumerate(zip(headers, widths), 1):
        c = ws.cell(row=8, column=ci, value=h)
        c.font = fnt(bold=True, color=LIGHT_GRAY, size=8)
        c.fill = fill(HEADER_BG); c.alignment = center()
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.row_dimensions[8].height = 15

    # Sub-cabecera de pesos
    ws.merge_cells("D7:G7")
    c = ws.cell(row=7, column=4, value="Pesos por fase")
    c.font = fnt(bold=True, color=LIGHT_GRAY, size=8)
    c.fill = fill(SECTION_BG); c.alignment = center()
    ws.row_dimensions[7].height = 13

    # ── Filas de sectores ──
    er = 9
    for sd in sector_data:
        alerta    = sd["alerta"]
        rec_w     = sd["rec_weight"]
        above_ema = sd["above_ema"]
        pct_ema   = sd["pct_ema"]

        # Color de fila base
        row_bg = "1F3244" if alerta else DARK_BG

        # Col 1: Nombre + alerta
        name_val = f"⚠ {sd['name']}" if alerta else sd["name"]
        c = ws.cell(row=er, column=1, value=name_val)
        c.font = fnt(bold=alerta, color="FFD700" if alerta else LIGHT_GRAY, size=8)
        c.fill = fill(row_bg); c.alignment = left(1)

        # Col 2: Ticker
        c = ws.cell(row=er, column=2, value=sd["ticker"])
        c.font = fnt(color=MID_GRAY, size=8, italic=True)
        c.fill = fill(row_bg); c.alignment = center()

        # Col 3: Precio
        c = ws.cell(row=er, column=3, value=fmt_price(sd["price"]))
        c.font = fnt(size=8); c.fill = fill(row_bg); c.alignment = center()

        # Cols 4-7: Pesos por fase (resaltar la fase activa)
        for fi, w in enumerate(sd["weights"]):
            col = fi + 4
            is_active = (fi == phase_idx)
            bg = phase_color if is_active else SECTION_BG
            tc = WHITE
            c = ws.cell(row=er, column=col, value=f"{w*100:.0f}%")
            c.font = fnt(bold=is_active, color=tc, size=8)
            c.fill = fill(bg); c.alignment = center()

        # Col 8: Peso recomendado actual (más grande y destacado)
        c = ws.cell(row=er, column=8, value=f"{rec_w*100:.0f}%")
        highlight = rec_w >= 0.15
        c.font = Font(bold=True, color=WHITE, size=10, name="Arial")
        c.fill = fill("1A6B3A" if highlight else SECTION_BG)
        c.alignment = center()

        # Cols 9-11: Retornos 1M, 3M, YTD
        for ri, rv in enumerate([sd["r1m"], sd["r3m"], sd["ytd"]]):
            c = ws.cell(row=er, column=9+ri, value=fmt_pct(rv))
            c.font = fnt(color=text_ret(rv), size=8, bold=(rv is not None and abs(rv) > 0.05))
            c.fill = fill(color_ret(rv)); c.alignment = center()

        # Col 12: vs EMA200 Semanal
        if above_ema is None:
            ema_txt = "N/A"
            ema_bg  = SECTION_BG
            ema_tc  = MID_GRAY
        elif above_ema:
            ema_txt = f"▲ +{pct_ema*100:.1f}%"
            ema_bg  = "1A5C2A"
            ema_tc  = "A9DFBF"
        else:
            ema_txt = f"▼ {pct_ema*100:.1f}%"
            ema_bg  = "7B241C"
            ema_tc  = "F1948A"

        c = ws.cell(row=er, column=12, value=ema_txt)
        c.font = fnt(bold=True, color=ema_tc, size=8)
        c.fill = fill(ema_bg); c.alignment = center()

        ws.row_dimensions[er].height = 15
        er += 1

    # ── Leyenda ──
    er += 1
    ws.merge_cells(f"A{er}:L{er}")
    c = ws.cell(row=er, column=1,
        value="  ⚠ Alerta: sector recomendado (peso ≥10%) con precio por debajo de EMA200 Semanal → revisar entrada  |  Peso Actual = peso recomendado en la fase detectada")
    c.font = fnt(italic=True, color=MID_GRAY, size=7)
    c.fill = fill(DARK_BG); c.alignment = left(1)
    ws.row_dimensions[er].height = 13

    # ── Tabla resumen de fases (abajo) ──
    er += 2
    ws.merge_cells(f"A{er}:L{er}")
    c = ws.cell(row=er, column=1, value="  REFERENCIA: Sectores favorecidos por fase")
    c.font = fnt(bold=True, color=LIGHT_GRAY, size=9)
    c.fill = fill(SECTION_BG); c.alignment = left(1)
    ws.row_dimensions[er].height = 14
    er += 1

    ref_data = [
        ("Expansión Temprana", "27AE60", "XLK (25%), XLY (20%), XLC (15%)  →  Consumidor y tecnología lideran el rebote"),
        ("Expansión Tardía",   "F39C12", "XLF (20%), XLI (20%), XLB (15%)  →  Ciclo maduro, inflación y tipos altos"),
        ("Recesión Temprana",  "E74C3C", "XLV (25%), XLP (20%), XLU (20%)  →  Defensivos, el mercado busca refugio"),
        ("Recesión Tardía",    "8E44AD", "XLE (15%), XLV (15%), XLRE (12%) →  Tipos bajan, energía y real estate se reactivan"),
    ]
    for phase, color, desc in ref_data:
        c1 = ws.cell(row=er, column=1, value=phase)
        c1.font = fnt(bold=True, size=8); c1.fill = fill(color); c1.alignment = left(1)
        ws.merge_cells(start_row=er, start_column=1, end_row=er, end_column=3)
        c2 = ws.cell(row=er, column=4, value=desc)
        c2.font = fnt(color=LIGHT_GRAY, size=8); c2.fill = fill(DARK_BG); c2.alignment = left(1)
        ws.merge_cells(start_row=er, start_column=4, end_row=er, end_column=12)
        ws.row_dimensions[er].height = 14
        er += 1

    ws.freeze_panes = "A9"

# ── MAIN ──────────────────────────────────────────────────────────────────────

def main():
    today_str = datetime.today().strftime("%d-%b-%y")
    output    = f"market_tracker_{datetime.today().strftime('%Y%m%d')}.xlsx"

    print("📡 Descargando precios de mercado...")
    all_tickers = [t for _, assets in SECTIONS for _, t in assets]
    # Añadir tickers SPI por si no están en SECTIONS
    spi_tickers = [t for _, t, *_ in SPI_SECTORS]
    all_tickers = list(set(all_tickers + spi_tickers))
    prices = get_prices(all_tickers)

    print("🔢 Calculando retornos...")
    market_rows = build_market_rows(prices)

    print("🌍 Descargando datos macro...")
    macro_data = build_macro_data()

    print("🔄 Calculando fase del ciclo SPI...")
    phase_idx, signals, sector_data = build_spi_data(prices, macro_data)
    print(f"   → Fase detectada: {PHASE_NAMES[phase_idx]}")

    print("📝 Generando Excel...")
    wb  = openpyxl.Workbook()
    ws1 = wb.active
    ws1.title = "Markets"
    write_market_sheet(ws1, market_rows, today_str)

    ws2 = wb.create_sheet("Macro")
    write_macro_sheet(ws2, macro_data, today_str)

    ws3 = wb.create_sheet("SPI")
    write_spi_sheet(ws3, phase_idx, signals, sector_data, today_str)

    wb.save(output)
    print(f"\n✅ Guardado: {output}")
    print(f"   Pestaña 1 → Markets  |  Pestaña 2 → Macro  |  Pestaña 3 → SPI ({PHASE_NAMES[phase_idx]})")

if __name__ == "__main__":
    main()
