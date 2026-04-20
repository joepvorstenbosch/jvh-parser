"""JVH Global — Offer Intake Tool v2.0"""
from __future__ import annotations
import io, re
from decimal import Decimal, InvalidOperation
from typing import Any, Dict, Optional
import pandas as pd
import streamlit as st

st.set_page_config(page_title="JVH Global — Offer Parser", page_icon="🥃", layout="wide")

st.markdown("""
<style>
  [data-testid="stAppViewContainer"] { background-color: #0d1b2a; color: #f0f0f0; }
  [data-testid="stHeader"] { background-color: #0d1b2a; }
  h1, h2, h3 { color: #f0a500 !important; font-family: Georgia, serif; }
  p, label, .stCaption { color: #cccccc !important; }
  .stTabs [data-baseweb="tab"] { color: #f0a500; background-color: #162233; border-radius: 6px 6px 0 0; padding: 8px 20px; }
  .stTabs [aria-selected="true"] { background-color: #f0a500 !important; color: #0d1b2a !important; font-weight: bold; }
  textarea, input[type="text"] { background-color: #162233 !important; color: #f0f0f0 !important; border: 1px solid #f0a500 !important; border-radius: 6px !important; }
  .stButton > button, .stDownloadButton > button { background-color: #f0a500 !important; color: #0d1b2a !important; font-weight: bold; border: none; border-radius: 6px; }
  .stButton > button:hover, .stDownloadButton > button:hover { background-color: #d4900a !important; }
  [data-testid="stMetric"] { background-color: #162233; border: 1px solid #f0a500; border-radius: 8px; padding: 12px; }
  [data-testid="stMetricValue"] { color: #f0a500 !important; font-size: 2rem !important; }
  [data-testid="stMetricLabel"] { color: #aaaaaa !important; }
  [data-testid="stDataFrame"] { border: 1px solid #f0a500; border-radius: 8px; }
  .stAlert { background-color: #162233 !important; border-left: 4px solid #f0a500 !important; color: #f0f0f0 !important; }
  .streamlit-expanderHeader { color: #f0a500 !important; background-color: #162233 !important; }
  [data-testid="stFileUploader"] { background-color: #162233; border: 1px dashed #f0a500; border-radius: 8px; padding: 10px; }
  .jvh-header { display: flex; align-items: center; gap: 18px; padding: 12px 0 20px 0; border-bottom: 2px solid #f0a500; margin-bottom: 24px; }
  .jvh-header img { height: 55px; }
  .jvh-header h1 { margin: 0; font-size: 1.8rem; color: #f0a500 !important; }
  .jvh-header p { margin: 2px 0 0 0; font-size: 0.85rem; color: #aaaaaa !important; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="jvh-header">
  <img src="https://www.jvh-global.com/assets/uploads/Rectangle-131@2x.png" alt="JVH Global">
  <div>
    <h1>Offer Parser</h1>
    <p>Plak een offerte of upload een bestand → download de gestructureerde Excel</p>
  </div>
</div>
""", unsafe_allow_html=True)

JVH_COLUMNS = [
    "Commodity","Product","GBX","Btls Case","Size CL","ABV %","RF NRF","ST",
    "Cases MOQ","# btls case","Freight cost","Cost per case",
    "Purchase Price - Bottle","Purchase Price - Case","Margin case","Margin, %",
    "Price per bottle","Price per Case","Currency","Incoterms","Leadtime",
    "Remark/BBD","Source Row","Parse Status","Review Flag","Review Notes",
]

COMMODITY_MAP = {
    "absolut":"Vodka","belvedere":"Vodka","ciroc":"Vodka","grey goose":"Vodka","smirnoff":"Vodka","finlandia":"Vodka",
    "rum":"Rum","bacardi":"Rum","captain morgan":"Rum","sailor jerry":"Rum","ron zacapa":"Rum","zacapa":"Rum",
    "gin":"Gin","tanqueray":"Gin","gin mare":"Gin","hendrick":"Gin","beefeater":"Gin","roku":"Gin","bombay":"Gin","bombay sapphire":"Gin",
    "tequila":"Tequila","jose cuervo":"Tequila","1800":"Tequila","sierra":"Tequila","don julio":"Tequila","olmeca":"Tequila","clase azul":"Tequila","espolon":"Tequila",
    "whisky":"Whisky","whiskey":"Whisky","jack daniel":"Whisky","jack daniels":"Whisky",
    "jim beam":"Whisky","jim beam white":"Whisky","jim beam cherry":"Whisky","jim beam original":"Whisky","jim beam apple":"Whisky","jim beam honey":"Whisky",
    "teachers":"Whisky","famous grouse":"Whisky","glenfiddich":"Whisky","glenlivet":"Whisky","hakushu":"Whisky","macallan":"Whisky","bowmore":"Whisky",
    "dewar":"Whisky","dewars":"Whisky","johnnie walker":"Whisky","hibiki":"Whisky","jameson":"Whisky","chivas":"Whisky",
    "grant's":"Whisky","grant s":"Whisky","lawson":"Whisky","auchentoshan":"Whisky","ballantines":"Whisky","highland park":"Whisky","royal brackla":"Whisky","aultmore":"Whisky",
    "glenmorangie":"Whisky","buchanan":"Whisky","grand old parr":"Whisky","old parr":"Whisky","seagram":"Whisky","pendleton":"Whisky",
    "crown royal":"Whisky","eagle rare":"Whisky","russell's reserve":"Whisky","russells reserve":"Whisky","tin cup":"Whisky","angels envy":"Whisky","angel's envy":"Whisky",
    "blantons":"Whisky","blanton":"Whisky","old grand-dad":"Whisky","old grand dad":"Whisky",
    "bourbon":"Whisky","canadian whiskey":"Whisky","blended american whiskey":"Whisky",
    "liqueur":"Liquor","liquor":"Liquor","aperol":"Liquor","jagermeister":"Liquor","licor 43":"Liquor","kahlua":"Liquor","malibu":"Liquor","grand marnier":"Liquor",
    "cognac":"Cognac","hennessy":"Cognac","martell":"Cognac","camus":"Cognac","remy martin":"Cognac","rémy martin":"Cognac",
    "champagne":"Champagne","veuve clicquot":"Champagne",
    "spritz":"RTD","wine":"Wines","sauvignon blanc":"Wines","jacobs creek":"Wines","oyster bay":"Wines","brancott":"Wines",
    "mini":"Miniatures (5cl)","minis":"Miniatures (5cl)","miniatures":"Miniatures (5cl)",
}

SECTION_TO_COMMODITY = {
    "RUM":"Rum","TEQUILA":"Tequila","VODKA":"Vodka","WHISKY":"Whisky","WHISKEY":"Whisky",
    "GIN":"Gin","COGNAC":"Cognac","CHAMPAGNE":"Champagne","WINE":"Wines","WINE (FCL)":"Wines","WINES":"Wines",
    "LIQUOR":"Liquor","BEERS":"Beers","SOFTDRINKS":"Softdrinks",
}

COLUMN_ALIASES = {
    "Lead Time":"Leadtime","Warehouse":"Incoterms","Coded":"ST","Cases Available (MOQ)":"Cases MOQ","Cases Available":"Cases MOQ",
    "RF/NRF":"RF NRF","REF/NRF":"RF NRF","producto":"Product","Producto":"Product",
    "btl/cs":"Btls Case","Btl/cs":"Btls Case","BTLS/CS":"Btls Case","CL":"Size CL","alc %":"ABV %","ABV%":"ABV %",
    "Price":"Price per bottle","price":"Price per bottle","cases":"Cases MOQ","CASES":"Cases MOQ","BRAND":"Product","SIZE LTR.":"Size LTR",
    "CAP":"RF NRF","STATUS":"ST","€/BTL":"Price per bottle","EUROS/CASE":"Price per Case","ETA":"Leadtime",
}

CURRENCY_SYMBOLS = {"€":"EUR","$":"USD","eur":"EUR","usd":"USD","euro":"EUR"}

def clean_text(v):
    if v is None: return ""
    t = str(v).replace("\u00a0"," ").replace("–","-").replace("—","-").replace("\u2019","'")
    return re.sub(r"[ \t]+"," ",t).strip()

def parse_decimal(v):
    t = clean_text(v).lower()
    if not t: return None
    for tok in ["eur","usd","euro","€","$","per bottle","per btl","per case","/btl","/cs"]: t = t.replace(tok,"")
    t = t.replace(" ","")
    if not t: return None
    if t.count(",") == 1 and t.count(".") >= 1: t = t.replace(".","").replace(",",".")
    elif t.count(",") == 1 and t.count(".") == 0: t = t.replace(",",".")
    try: return Decimal(t)
    except InvalidOperation: return None

def format_money(v, currency):
    if v is None: return ""
    return f"{currency} {v:.2f}" if currency in ("EUR","USD") else f"{v:.2f}"

def detect_currency(*values):
    joined = " ".join(clean_text(v).lower() for v in values if clean_text(v))
    for sym, code in CURRENCY_SYMBOLS.items():
        if sym in joined: return code
    return ""

def to_int(v):
    t = clean_text(v)
    if not t: return None
    m = re.search(r"\d+", t)
    return int(m.group()) if m else None

def to_float(v):
    d = parse_decimal(v)
    return float(d) if d is not None else None

def infer_commodity(product, size_cl=None):
    p = clean_text(product).lower()
    if size_cl == 5 or "mini" in p or re.search(r"\b5\s?cl\b", p): return "Miniatures (5cl)"
    for key, commodity in COMMODITY_MAP.items():
        if key in p: return commodity
    return ""

def standardize_incoterms(v):
    t = clean_text(v)
    t = re.sub(r"(?i)\bexw\b","Exworks",t)
    t = re.sub(r"(?i)\bex\b","Exworks",t)
    return t

def standardize_leadtime(v):
    text = clean_text(v).replace("Lead time","").replace("Leadtime","").strip()
    t = text.lower().strip()
    if not t: return ""
    if re.search(r"(?i)\bon\s*floor\b|\bex\s*stock\b|\bin\s*stock\b|\bstock\b|\bready\b", t): return "On floor"
    # "Third week of May" -> "Week 3 May"
    ordinals = {"first":"1","second":"2","third":"3","fourth":"4"}
    m = re.match(r"(?i)^(first|second|third|fourth)\s+week\s+of\s+([a-z]+)$", t)
    if m: return f"Week {ordinals.get(m.group(1).lower(),'?')} {m.group(2).capitalize()}"
    m = re.match(r"(?i)^(mid|end|early|begin)\s+([a-z]+)$", t)
    if m: return m.group(1).capitalize()+" "+m.group(2).capitalize()
    m = re.match(r"(\d+)\s*-\s*(\d+)\s*days?", t)
    if m: return f"{m.group(1)}-{m.group(2)} Days"
    m = re.match(r"(\d+)\s*days?", t)
    if m: return f"{m.group(1)} Days"
    m = re.match(r"(\d+)\s*-\s*(\d+)\s*weeks?", t)
    if m:
        lo,hi = int(m.group(1)),int(m.group(2)); return f"{lo*5}-{hi*5} Days"
    m = re.match(r"(\d+)\s*weeks?", t)
    if m:
        w = int(m.group(1)); return f"{w*5}-{(w+1)*5} Days"
    return text

def standardize_rf(v):
    t = clean_text(v).upper().replace(".","")
    if not t: return "REF"
    if t in {"RF","REF","REFILLABLE"}: return "REF"
    if t in {"NRF","NON-REF","NON REF","NONREF"}: return "NRF"
    return t

def ensure_jvh_columns(df):
    for col in JVH_COLUMNS:
        if col not in df.columns: df[col] = ""
    return df[JVH_COLUMNS]

def build_output_row(data):
    row = {col:"" for col in JVH_COLUMNS}
    row.update(data)
    row["# btls case"] = row.get("Btls Case","")
    row["Price per bottle"] = row.get("Purchase Price - Bottle","")
    row["Price per Case"] = row.get("Purchase Price - Case","")
    return row

def parse_column_blob(blob):
    """Parse samengesmolten kolommen-offerte: qty|product|prijs|warehouse|leadtime|status."""
    from decimal import Decimal as D
    blob = re.sub(r"(?i)Quantity\s*Product\s*Price\s*Warehouse\s*Lead\s*Time\s*Coded\s*","",blob)
    blob = re.sub(r"(?i)Description\s*QTY\s*BOTT[A-Z\s]*Lead\s*time\s*","",blob)
    pattern = re.compile(
        r"((?:FCL|FTL|\d[\d,]*\s*(?:btls?|cs|cases?)).*?)"
        r"((?:Euro|\u20ac|\$|USD)\s*[\d]+[,.]?[\d]*\s*/(?:btl|cs))"
        r"(.*?)"
        r"(?=(?:FCL|FTL)(?:[A-Za-z]|\s)|(?<!\d)\d[\d,]*\s+(?:btls?|cs|cases?)|\Z)",
        re.IGNORECASE | re.DOTALL
    )
    rows = []
    for pre, price_raw, post in pattern.findall(blob):
        pre = pre.strip()
        # Strip T1/T2 digit prefix left over from previous record
        # Pattern: single digit 1 or 2 before a large number = T1 or T2 leftover
        pre = re.sub(r"^[12](?=\d{3,})", "", pre)
        pre = re.sub(r"^(T[12]|Coded)\s*","",pre,flags=re.I)
        qty_match = re.match(r"(?i)^(FCL|FTL)\s*(.*)", pre)
        num_match = re.match(r"(?i)^([\d,]+)\s*(btls?|cs|cases?)\s*(.*)", pre)
        if qty_match:
            qty_type, qty, product_raw = "FTL", 0, qty_match.group(2).strip()
        elif num_match:
            qty = int(num_match.group(1).replace(",",""))
            qty_type = "BTLS" if "btl" in num_match.group(2).lower() else "CS"
            product_raw = num_match.group(3).strip()
        else:
            qty_type, qty, product_raw = "FTL", 0, pre

        btls_case = size_cl = abv = None
        # "6/ 50/40" or "6/70/40" -> btls=6, size=70cl, abv=40%
        sm = re.search(r"\b(\d+)/\s*(\d+)/\s*(\d+)\b", product_raw)
        if sm:
            btls_case = int(sm.group(1))
            size_cl = int(sm.group(2))
            abv = float(sm.group(3))
            product_raw = (product_raw[:sm.start()] + " " + product_raw[sm.end():]).strip()

        # "100cl 40%" or "70cl" alone - size without btls_case
        if size_cl is None:
            cm = re.search(r"\b(\d+)cl\b", product_raw, re.I)
            if cm: size_cl = int(cm.group(1))

        # ABV%
        if abv is None:
            am = re.search(r"\b(\d{1,2}(?:[.,]\d+)?)%\b", product_raw)
            if am: abv = float(am.group(1).replace(",","."))

        gbx = "GBX" if re.search(r"(?i)\b(gbx|ngbx)\b", product_raw) else ""
        rf_nrf = "NRF" if re.search(r"(?i)\bNRF\b|\bnon.?ref\b", product_raw) else "REF"

        # Clean product name
        product = re.sub(r"(?i)\b(ref|ngbx|gbx|nrf)\b","",product_raw)
        product = re.sub(r"\(glass bottle\)","",product,flags=re.I)
        product = re.sub(r"\b\d{1,2}(?:[.,]\d+)?%","",product)
        product = re.sub(r"\b\d+cl\b","",product,flags=re.I)  # strip "100cl" from name
        product = re.sub(r"\b40%\b|\b35%\b","",product)  # strip ABV from name
        # Strip warehouse/location names that end up in product
        product = re.sub(r"(?i)\s*/\s*newcorp.*$","",product)
        product = re.sub(r"(?i)\s*/\s*loendersloot.*$","",product)
        product = re.sub(r"(?i)\bon\s+the\s+floor\b","",product)
        product = re.sub(r"(?i)\bthe\s+floor\b","",product)
        product = re.sub(r"(?i)\b(week\s+of\s+may|third\s+week.*$)","",product)
        product = re.sub(r"\s+"," ",product).strip(" .,-/")

        # Price
        pv = re.search(r"[\d]+[,.]?[\d]*", price_raw)
        price_num = pv.group().replace(",",".") if pv else None
        currency = "USD" if re.search(r"(?i)\$|USD", price_raw) else "EUR"

        # Post: warehouse, leadtime, status
        post = re.sub(r"(?i)(?<!\s)(On the floor|On floor|Coded|T[12])", r" \1", post)
        inc = re.search(r"(?i)\b(DAP\s+[A-Za-z]+(?:\s+[A-Za-z]+)?|Exw(?:orks)?\s+[A-Za-z]+(?:\s+[A-Za-z]+)?|EXW\s+[A-Za-z]+)", post)
        incoterms = re.sub(r"(?i)\bExw\b","Exworks",inc.group(1)) if inc else ""
        # Strip leadtime/week info that ended up in incoterms
        incoterms = re.sub(r"(?i)\s+(third|second|first|fourth|week|of|on|floor|coded).*$","",incoterms).strip()
        incoterms = re.sub(r"(?i)\s+On$","",incoterms).strip()

        st_m = re.search(r"(?i)\bT[12]\b", post)
        st = st_m.group(0).upper() if st_m else ""

        lm = re.search(r"(?i)(on\s+(?:the\s+)?floor|\d+\s*-\s*\d+\s*(?:weeks?|days?)|\d+\s*(?:weeks?|days?)|(?:first|second|third|fourth)\s+week\s+of\s+[a-z]+|mid\s+[a-z]+)", post)
        if lm:
            lt = lm.group(1).lower().strip()
            if "floor" in lt: leadtime = "On floor"
            elif m2 := re.match(r"(\d+)\s*-\s*(\d+)\s*weeks?", lt):
                leadtime = f"{int(m2.group(1))*5}-{int(m2.group(2))*5} Days"
            elif m2 := re.match(r"(\d+)\s*weeks?", lt):
                w=int(m2.group(1)); leadtime = f"{w*5}-{(w+1)*5} Days"
            elif m2 := re.match(r"(\d+)\s*-\s*(\d+)\s*days?", lt):
                leadtime = f"{m2.group(1)}-{m2.group(2)} Days"
            elif re.match(r"(?i)(first|second|third|fourth)\s+week", lt):
                ord_map={"first":"1","second":"2","third":"3","fourth":"4"}
                pts = lt.split(); leadtime = f"Week {ord_map.get(pts[0],'?')} {pts[-1].capitalize()}"
            else: leadtime = lm.group(1)
        else: leadtime = ""

        cases_moq = "FTL" if qty_type == "FTL" else (qty // btls_case if qty_type == "BTLS" and btls_case else qty)
        price_d = D(price_num) if price_num else None
        btl_price = f"{currency} {price_d:.2f}" if price_d else ""
        case_price = f"{currency} {price_d * D(str(btls_case)):.2f}" if price_d and btls_case else ""
        remark = "CODED" if re.search(r"(?i)\bcoded\b", post) else ""
        infer = infer_commodity(product, size_cl)

        missing = [x for x in [
            "Missing Btls Case" if not btls_case else "",
            "Missing Size CL" if not size_cl else "",
            "Missing ABV % (handmatig invullen)" if not abv else "",
            "Missing Incoterms" if not incoterms else "",
            "Missing Leadtime" if not leadtime else "",
        ] if x]

        rows.append(build_output_row({
            "Commodity": infer, "Product": product, "GBX": gbx,
            "Btls Case": btls_case, "Size CL": size_cl, "ABV %": abv,
            "RF NRF": rf_nrf, "ST": st, "Cases MOQ": cases_moq,
            "Purchase Price - Bottle": btl_price, "Purchase Price - Case": case_price,
            "Currency": currency, "Incoterms": incoterms, "Leadtime": leadtime,
            "Remark/BBD": remark, "Source Row": "blob",
            "Parse Status": "REVIEW" if missing else "OK",
            "Review Flag": "YES" if missing else "NO",
            "Review Notes": "; ".join(missing),
        }))
    return ensure_jvh_columns(pd.DataFrame(rows)) if rows else pd.DataFrame()


def detect_and_split_blob(text):
    """Detecteer blob-offerte en splits naar losse regels voor de normale parser."""
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    def is_blob(l):
        has_price = l.count("\u20ac") >= 2 or len(re.findall(r"(?i)\beuro\b", l)) >= 2
        return len(l) > 150 and has_price
    blob_lines = [l for l in lines if is_blob(l)]
    if not blob_lines: return text
    result_lines = [l for l in lines if l not in blob_lines]
    for blob in blob_lines:
        blob = re.sub(r"(?i)Description\s*QTY\s*BOTT[A-Z\s]*Lead\s*time\s*","",blob)
        blob = re.sub(r"(T[12]|Coded)(FCL|FTL)", r"\1\n\2", blob, flags=re.I)
        blob = re.sub(r"(T[12]|Coded)(\d+\s+(?:btls?|cs|cases?))", r"\1\n\2", blob, flags=re.I)
        blob = re.sub(r"(?i)(floor)(FCL|FTL)", r"\1\n\2", blob)
        blob = re.sub(r"(?i)(floor)(\d+\s+(?:btls?|cs|cases?))", r"\1\n\2", blob)
        for part in blob.splitlines():
            part = part.strip()
            if not part or len(part) < 8: continue
            part = re.sub(r"(?i)^(FCL|FTL)([A-Za-z])", r"\1 \2", part)
            part = re.sub(r"(?i)^FCL\b","FTL",part)
            for kw in ["DAP ","Exw ","Exworks ","CFR ","On the floor","On floor","Coded","T1 ","T2 "]:
                part = re.sub(rf"(?<!\s)({re.escape(kw.strip())})", r" \1", part)
            part = re.sub(r"\b(\d+)/(\d+)/(\d+)\b", lambda m: f"{m.group(1)}x{m.group(2)}cl {m.group(3)}%", part)
            part = re.sub(r"\(glass bottle\)","",part,flags=re.I)
            part = re.sub(r"(\d+(?:\.\d+)?)[Ll]\s*[xX]\s*(\d+)", lambda m: f"{m.group(2)}x{m.group(1)}L", part)
            part = re.sub(r"(?i)\bON THE FLOOR\s*LOEND\.?","On floor Exworks Loendersloot",part)
            part = re.sub(r"(?i)\bON THE FLOOR\b","On floor",part)
            part = re.sub(r"(?i)\bExw\b","Exworks",part)
            part = re.sub(r"(?i)\bEuros?\s*([\d]+[,.]?[\d]*)\s*/(btl|cs)\b",
                          lambda m: f"@ EUR {m.group(1).replace(',','.')} /{m.group(2)}", part)
            part = re.sub(r"\u20ac\s*([\d,]+(?:\.\d+)?)", lambda m: f"@ EUR {m.group(1).replace(',','.')} /btl", part)
            result_lines.append(part)
    return "\n".join(result_lines)


def preprocess_text(text):
    for old, new in {"RF.":"RF","/cs.":"/cs","/btl.":"/btl","/btl.,":"/btl,","/cs.,":"/cs,"," per bottle":" /btl"," per case":" /cs"," per cs":" /cs"," per btl":" /btl"}.items():
        text = text.replace(old, new)
    cleaned = []
    for raw in text.splitlines():
        line = raw.strip()
        if not line: cleaned.append(""); continue
        line = re.sub(r"^[-*•]\s*","",line)
        line = re.sub(r"(?<=\d),(?=\d{3}\b)","",line)
        line = line.replace("–","-").replace("—","-")
        line = re.sub(r"(?i)^FCL\b","FTL",line)
        line = re.sub(r"\b(\d+)/(\d+)/(\d+)\b", lambda m: f"{m.group(1)}x{m.group(2)}cl {m.group(3)}%", line)
        line = re.sub(r"\(glass bottle\)","",line,flags=re.I)
        line = re.sub(r"\((\d+(?:\.\d+)?)[Ll]\s*[xX]\s*(\d+)\)", lambda m: f"{m.group(2)}x{m.group(1)}L", line)
        line = re.sub(r"\((\d+)[Cc][Ll]\s*[xX]\s*(\d+)\)", lambda m: f"{m.group(2)}x{m.group(1)}cl", line)
        line = re.sub(r"(?i)\bQty\s*:\s*","",line)
        # Normalize "Euro 7,49/btl" or "Euro3,80/btl" -> "@ EUR 7.49 /btl"
        def fix_euro_price(line):
            def repl(m):
                price = m.group(1).replace(',','.')
                unit = m.group(2)
                return f"@ EUR {price} /{unit}"
            # With or without space after Euro, with comma or dot decimal
            line = re.sub(r"(?i)\bEuros?\s*([\d]+[,.][\d]+)\s*/(btl|cs)\b", repl, line)
            line = re.sub(r"(?i)\bEuros?\s+([\d]+)\s*/(btl|cs)\b",
                          lambda m: f"@ EUR {m.group(1)} /{m.group(2)}", line)
            return line
        if not re.search(r"@\s*(EUR|USD)", line):
            line = fix_euro_price(line)
        # Fix "6/ 50/40" (space after slash) -> "6/50/40"
        line = re.sub(r"(\d+)/\s+(\d+)/\s*(\d+)", r"\1/\2/\3", line)
        line = re.sub(r"(?i)\bPrice\s*:\s*€\s*(\d+(?:[.,]\d+)?)\s*/(btl|cs)\b",r"@ EUR \1 /\2",line)
        line = re.sub(r"(?i)\bPrice\s*:\s*\$\s*(\d+(?:[.,]\d+)?)\s*/(btl|cs)\b",r"@ USD \1 /\2",line)
        line = re.sub(r"(?i)\bPrice\s*:\s*(USD|EUR)\s*(\d+(?:[.,]\d+)?)\s*/(btl|cs)\b",r"@ \1 \2 /\3",line)
        # Only normalize bare € or $ if no @ already present
        if not re.search(r"@\s*(EUR|USD|€|\$)", line):
            line = re.sub(r"€\s*(\d+(?:[.,]\d+)?)\s*/(btl|cs)\b",r"@ EUR \1 /\2",line)
            line = re.sub(r"\$\s*(\d+(?:[.,]\d+)?)\s*/(btl|cs)\b",r"@ USD \1 /\2",line)
        line = re.sub(r"(?i)\bex-([A-Za-z]+)",r"ex \1",line)
        line = re.sub(r"(?i)\bDuty\s*Status\s*:\s*","",line)
        if not re.search(r"@\s*(EUR|USD)", line):
            line = re.sub(r"(?i)\beuro\s+(\d)",r"@ EUR \1",line)
            line = re.sub(r"(?i)\beuros\s+(\d)",r"@ EUR \1",line)
        # Fix any remaining "@ EUR 7,49" -> "@ EUR 7.49" (comma decimal after @)
        line = re.sub(r"(@\s*(?:EUR|USD)\s+)(\d+),(\d+)", r"\g<1>\2.\3", line)
        line = re.sub(r"(?i)\b(\d+(?:[.,]\d+)?)\s*(USD|EUR)\s*(?:per case|/cs)?\s*$",r"@ \2 \1 /cs",line)
        line = re.sub(r"(?i)\b(USD|EUR)\s+(\d+(?:[.,]\d+)?)\s+per case\b",r"@ \1 \2 /cs",line)
        line = re.sub(r"(?i)\b(USD|EUR)\s+(\d+(?:[.,]\d+)?)\s*/(cs|btl)\b",r"@ \1 \2 /\3",line)
        # FTL zonder qty en zonder prijs -> voeg placeholder toe
        if re.match(r"(?i)^FTL\b",line) and not re.search(r"(?i)\d+\s*(cases|case|cs|bottles|bottle|btls)\b",line) and not re.search(r"@",line):
            line = line + " 1 cs FTL_LINE"
        cleaned.append(line)
    return "\n".join(cleaned)

def parse_vertical_offer(text):
    """Parse vertikale tabel-offerte: elke waarde op eigen regel (DESCRIPTION/BTL/ML/ABV/...)."""
    from decimal import Decimal as D
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    # Skip header block
    header_markers = {"DESCRIPTION","BTL/CTN","ML/BTL","ABV","BOTTLE","VARIANT",
                      "EU STATUS","CARTON","EURO EXW LOENDERSLOOT","EURO EXW",
                      "PRICE BTL","PRICE CARTON","PRICE/BTL","PRICE/CARTON"}
    header_end = 0
    for i, line in enumerate(lines):
        if any(line.upper().startswith(h) for h in header_markers):
            header_end = i + 1
        elif header_end > 0:
            break
    data_lines = lines[header_end:]
    # Detect incoterms from header area
    incoterms = ""
    for line in lines[:header_end]:
        inc = re.search(r"(?i)\b(EXW?\s+[A-Za-z]+(?:\s+[A-Za-z]+)?|DAP\s+[A-Za-z]+)", line)
        if inc:
            incoterms = re.sub(r"(?i)\bExw\b","Exworks",inc.group(1))
            break
    if not incoterms: incoterms = "Exworks Loendersloot"
    # Fix kapitalisatie
    incoterms = re.sub(r"\bLOENDERSLOOT\b","Loendersloot",incoterms)
    incoterms = re.sub(r"\bRIGA\b","Riga",incoterms)
    incoterms = re.sub(r"\bSINGAPORE\b","Singapore",incoterms)
    # Each record = 9 fields: name, btls, ml, abv, rf_nrf, variant, st, price_carton, price_bottle
    FIELDS = 9
    rows = []
    i = 0
    while i < len(data_lines):
        line = data_lines[i]
        is_product = (
            not re.match(r"^[\d€$.,]+", line) and
            not re.match(r"^(T[12]|REF|NRF|NGB|GBX|GB)\s*$", line, re.I)
        )
        if is_product and i + FIELDS - 1 < len(data_lines):
            rec = data_lines[i:i+FIELDS]
            name, btls_s, ml_s, abv_s, rf_s, variant_s, st_s, price_ctn_s, price_btl_s = rec
            btls_case = int(btls_s) if re.match(r"^\d+$", btls_s) else None
            ml = int(float(ml_s)) if re.match(r"^[\d.]+$", ml_s) else None
            size_cl = ml // 10 if ml else None
            abv = float(abv_s.replace(",",".")) if re.match(r"^[\d.,]+$", abv_s) else None
            rf_nrf = "NRF" if "NRF" in rf_s.upper() else "REF"
            gbx = "GBX" if re.search(r"(?i)\bGBX\b", variant_s) else ""
            st = st_s.upper() if re.match(r"^T[12]$", st_s, re.I) else ""
            def parse_price(s):
                m = re.search(r"[\d]+[.,][\d]+|[\d]+", s.replace(" ",""))
                return m.group().replace(",",".") if m else None
            p_ctn = parse_price(price_ctn_s)
            p_btl = parse_price(price_btl_s)
            currency = "USD" if "$" in price_btl_s else "EUR"
            ctn_d = D(p_ctn) if p_ctn else None
            btl_d = D(p_btl) if p_btl else None
            # Cross-check: if both present, verify consistency
            if ctn_d and btl_d and btls_case:
                expected_ctn = btl_d * btls_case
                if abs(ctn_d - expected_ctn) > D("0.10"):
                    pass  # accept as-is, prices per bottle and carton may differ
            product = re.sub(r"\s+"," ", name).strip()
            commodity = infer_commodity(product, size_cl)
            missing = [x for x in [
                "Missing Btls Case" if not btls_case else "",
                "Missing Size CL" if not size_cl else "",
                "Missing ABV % (handmatig invullen)" if not abv else "",
            ] if x]
            rows.append(build_output_row({
                "Commodity": commodity, "Product": product, "GBX": gbx,
                "Btls Case": btls_case, "Size CL": size_cl, "ABV %": abv,
                "RF NRF": rf_nrf, "ST": st, "Cases MOQ": "",
                "Purchase Price - Bottle": f"{currency} {btl_d:.2f}" if btl_d else "",
                "Purchase Price - Case": f"{currency} {ctn_d:.2f}" if ctn_d else "",
                "Currency": currency, "Incoterms": incoterms, "Leadtime": "",
                "Remark/BBD": "", "Source Row": str(i + header_end + 1),
                "Parse Status": "REVIEW" if missing else "OK",
                "Review Flag": "YES" if missing else "NO",
                "Review Notes": "; ".join(missing),
            }))
            i += FIELDS
        else:
            i += 1
    return ensure_jvh_columns(pd.DataFrame(rows)) if rows else pd.DataFrame()

def is_vertical_offer(text):
    """Detecteer of tekst een verticale tabel-offerte is."""
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    markers = sum(1 for l in lines[:15] if l.upper() in
                  {"DESCRIPTION","BTL/CTN","ML/BTL","ABV","EU STATUS","VARIANT"})
    return markers >= 3

def parse_offer_text(text):
    # Detecteer verticale tabel-offerte
    if is_vertical_offer(text):
        return parse_vertical_offer(text)
    # Detecteer kolommen-blob (Euro-prijs formaat zonder spaties)
    lines_raw = [l.strip() for l in text.splitlines() if l.strip()]
    def is_col_blob(l):
        has_euro = len(re.findall(r"(?i)\beuro\b", l)) >= 2
        return len(l) > 150 and has_euro
    col_blobs = [l for l in lines_raw if is_col_blob(l)]
    if col_blobs:
        all_rows = []
        for blob in col_blobs:
            df_blob = parse_column_blob(blob)
            if not df_blob.empty:
                all_rows.append(df_blob)
        non_blob = [l for l in lines_raw if not is_col_blob(l)]
        if non_blob:
            df_rest = _parse_offer_text_inner("\n".join(non_blob))
            if not df_rest.empty:
                all_rows.append(df_rest)
        if all_rows:
            return ensure_jvh_columns(pd.concat(all_rows, ignore_index=True))
        return pd.DataFrame()
    return _parse_offer_text_inner(text)

def _parse_offer_text_inner(text):
    text = detect_and_split_blob(text)
    text = preprocess_text(text)
    rows = []
    current_section = ""
    default_coded = False
    trailing_bbd = ""
    trailing_moq = None
    global_incoterms = ""
    global_leadtime = ""

    for raw in text.splitlines():
        line = clean_text(raw)
        if not line: continue
        if not re.search(r"(?i)@|\d+\s*(cs|cases|btls)\b", line):
            inc = re.search(r"(?i)\b(EX(?:W| |\s)\s*[A-Za-z]+(?:\s+[A-Za-z]+)?|DAP\s+[A-Za-z]+(?:\s+[A-Za-z]+)?)\b", line)
            if inc and not global_incoterms: global_incoterms = standardize_incoterms(inc.group(1))
            lead = re.search(r"(?i)(on\s*floor|stock|ready|\d+\s*-\s*\d+\s*(?:weeks?|days?)|\d+\s*(?:weeks?|days?))", line)
            if lead and not global_leadtime: global_leadtime = standardize_leadtime(lead.group(1))

    for line_number, raw_line in enumerate(text.splitlines(), start=1):
        line = clean_text(raw_line)
        if not line: continue
        upper = line.upper()
        if upper.rstrip(":") in SECTION_TO_COMMODITY:
            current_section = SECTION_TO_COMMODITY[upper.rstrip(":")]; continue
        if upper.startswith("BBD:") or "MOQ:" in upper:
            m = re.search(r"(?i)bbd:\s*([^-]+(?:[-/][^-]+)?)", line)
            if m: trailing_bbd = clean_text(m.group(1))
            m = re.search(r"(?i)moq:\s*([\d,]+)\s*cs", line)
            if m: trailing_moq = int(m.group(1).replace(",",""))
            continue
        if "CODED" in upper and "ALL ITEMS" in upper:
            default_coded = True; continue

        is_ftl = bool(re.search(r"(?i)\bFTL\b", line)) or "FTL_LINE" in line
        qty_match = re.search(r"(?i)([\d,]+)\s*(cases|case|cs|bottles|bottle|btls)\b", line)
        if not qty_match and not is_ftl: continue
        if qty_match:
            qty_raw = qty_match.group(1).replace(",","")
            qty = 0 if qty_raw == "FTL" else int(qty_raw)
            qty_unit = "BTLS" if qty_match.group(2).lower() in {"bottles","bottle","btls"} else "CS"
        else:
            qty = 0; qty_unit = "CS"

        price_match = re.search(r"(?i)@\s*(EUR|USD|€|\$)?\s*([0-9]+(?:[.,][0-9]+)?)\s*(?:/(btl|cs)|per\s+(bottle|case))?", line)
        if not price_match: continue
        curr_token = clean_text(price_match.group(1)).lower()
        currency = "USD" if curr_token in {"usd","$"} else "EUR" if curr_token in {"eur","€","euro","euros"} else ""
        raw_price = parse_decimal(price_match.group(2))
        price_type = (price_match.group(3) or price_match.group(4) or "").lower()

        btls_case = None; size_cl = None
        m = re.search(r"(?i)(\d+)x(\d+(?:\.\d+)?)l\b", line)
        if m:
            btls_case = int(m.group(1)); size_cl = int(Decimal(m.group(2)) * Decimal("100"))
        else:
            m = re.search(r"(?i)(\d+)x(\d+(?:\.\d+)?)ml\b", line)
            if m:
                btls_case = int(m.group(1)); size_cl = int(Decimal(m.group(2)) / Decimal("10"))
            else:
                m = re.search(r"(?i)(\d+)x(\d+)cl\b", line)
                if m:
                    btls_case = int(m.group(1)); size_cl = int(m.group(2))

        abv_match = re.search(r"(?i)\b(\d{1,2}(?:[.,]\d+)?)%\b", line)
        abv = float(abv_match.group(1).replace(",",".")) if abv_match else None
        rf_nrf = "NRF" if re.search(r"(?i)\bNRF\b|\bnon.?ref\b", line) else "REF"
        st_match = re.search(r"(?i)\bT[12]\b", line)
        st_status = st_match.group(0).upper() if st_match else ""

        parts_pipe = [clean_text(p) for p in line.split("|")]
        incoterms = standardize_incoterms(parts_pipe[-2]) if len(parts_pipe) >= 3 else ""
        leadtime = standardize_leadtime(parts_pipe[-1]) if len(parts_pipe) >= 2 else ""
        if not incoterms:
            inc_match = re.search(r"(?i)\b(EX(?:W| |\s)\s*[A-Za-z]+(?:\s+[A-Za-z]+)?|DAP\s+[A-Za-z]+(?:\s+[A-Za-z]+)?|CFR\s+[A-Za-z]+|CNF\s+[A-Za-z]+)\b", line)
            if inc_match: incoterms = standardize_incoterms(inc_match.group(1))
        if not incoterms and global_incoterms: incoterms = global_incoterms
        if not leadtime:
            lead_match = re.search(r"(?i)(on\s*floor|stock|ready|\d+\s*-\s*\d+\s*(?:weeks?|days?)|\d+\s*(?:weeks?|days?)|mid\s+[a-z]+|end\s+[a-z]+|early\s+[a-z]+|(?:first|second|third|fourth)\s+week\s+of\s+[a-z]+)", line)
            if lead_match: leadtime = standardize_leadtime(lead_match.group(1))
        if not leadtime and global_leadtime: leadtime = global_leadtime

        product = line
        for p in [
            r"(?i)\bFTL\b", r"\bFTL_LINE\b", r"\b1\s+cs\b",
            r"(?i)\b[\d,]+\s*(cases|case|cs|bottles|bottle|btls)\b",
            r"(?i)@\s*(EUR|USD|€|\$)?\s*[0-9]+(?:[.,][0-9]+)?\s*(?:/(?:btl|cs)|per\s+(?:bottle|case))?",
            r"(?i)\b(EX(?:W| |\s)\s*[A-Za-z]+(?:\s+[A-Za-z]+)?|DAP\s+[A-Za-z]+(?:\s+[A-Za-z]+)?|CFR\s+[A-Za-z]+|CNF\s+[A-Za-z]+)\b",
            r"(?i)\b(on\s*floor|stock|ready|\d+\s*-\s*\d+\s*(?:weeks?|days?)|\d+\s*(?:weeks?|days?)|mid\s+[a-z]+|(?:first|second|third|fourth)\s+week\s+of\s+[a-z]+)\b",
            r"(?i)\bNRF\b|\bREF\b|\bRF\b|\bT1\b|\bT2\b|\(coded\)|coded",
            r"(?i)\b\d+x\d+(?:\.\d+)?l\b|\b\d+x\d+cl\b|\b\d+x\d+ml\b",
            r"(?i)\bDuty\s*Status\s*:?\s*T[12]\b",
            r"(?i)\bPrice\s*:", r"(?i)\bQty\s*:",
            r"\b\d{1,2}(?:[.,]\d+)?%\b",
            r"\(.*?\)",
        ]:
            product = re.sub(p,"",product)
        product = product.replace("|"," ").replace("--"," ")
        product = re.sub(r"(?i)\+?\s*gb\b"," GBX",product)
        gbx = "GBX" if re.search(r"(?i)\bgbx\b|\bcradle\b|\bgb\b", product) else ""
        product = re.sub(r"(?i)\bgbx\b|\bcradle\b","",product)
        product = re.sub(r"[@,]"," ",product)
        product = re.sub(r"\s+-\s*$|\s*-\s+"," ",product)
        product = re.sub(r"\s+"," ",product).strip(" .,-@")
        commodity = current_section or infer_commodity(product, size_cl)

        bottle_price = case_price = None
        if price_type in {"bottle","btl"}:
            bottle_price = raw_price
            if raw_price is not None and btls_case: case_price = raw_price * Decimal(str(btls_case))
        else:
            case_price = raw_price
            if raw_price is not None and btls_case: bottle_price = raw_price / Decimal(str(btls_case))

        cases_moq = "FTL" if (is_ftl and qty == 0) else qty
        review_notes = []
        if qty_unit == "BTLS" and qty > 0:
            review_notes.append("Hoeveelheid in flessen — omgerekend naar dozen")
            if btls_case: cases_moq = qty // btls_case
        if not commodity: review_notes.append("Missing commodity")
        if btls_case is None: review_notes.append("Missing Btls Case")
        if size_cl is None: review_notes.append("Missing Size CL")
        if abv is None: review_notes.append("Missing ABV % (handmatig invullen)")
        if not currency: review_notes.append("Missing currency — controleer offerte")
        if not leadtime: review_notes.append("Missing Leadtime")
        if not incoterms: review_notes.append("Missing Incoterms")

        remark_parts = []
        if default_coded or "coded" in line.lower(): remark_parts.append("CODED")
        if trailing_bbd: remark_parts.append(f"BBD: {trailing_bbd}")
        if trailing_moq: remark_parts.append(f"MOQ: {trailing_moq} cs")

        rows.append(build_output_row({
            "Commodity":commodity,"Product":product,"GBX":gbx,
            "Btls Case":btls_case,"Size CL":size_cl,"ABV %":abv,
            "RF NRF":rf_nrf,"ST":st_status,"Cases MOQ":cases_moq,
            "Purchase Price - Bottle":format_money(bottle_price,currency),
            "Purchase Price - Case":format_money(case_price,currency),
            "Currency":currency,"Incoterms":incoterms,"Leadtime":leadtime,
            "Remark/BBD":" | ".join(remark_parts),"Source Row":str(line_number),
            "Parse Status":"REVIEW" if any(n.startswith("Missing") for n in review_notes) else "OK",
            "Review Flag":"YES" if review_notes else "NO",
            "Review Notes":"; ".join(review_notes),
        }))
    return ensure_jvh_columns(pd.DataFrame(rows))

def to_excel_bytes(df):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, sheet_name="JVH Master", index=False)
        df[df["Review Flag"]=="YES"].copy().to_excel(w, sheet_name="Review", index=False)
    buf.seek(0)
    return buf.getvalue()

col_in, col_out = st.columns([1, 1])

with col_in:
    st.subheader("📥 Input")
    tab_text, tab_file = st.tabs(["Tekst plakken", "Bestand uploaden"])
    with tab_text:
        pasted = st.text_area("Plak hier de offertetekst (email, WhatsApp, etc.)", height=400, placeholder="Plak de volledige offertetekst hier...")
        supplier = st.text_input("Leverancier (optioneel)", placeholder="bijv. Diageo / Pernod / ...")
    with tab_file:
        uploaded = st.file_uploader("Upload Excel of CSV", type=["xlsx","xls","csv"])
        supplier_f = st.text_input("Leverancier (optioneel) ", placeholder="bijv. Diageo / Pernod / ...")

with col_out:
    st.subheader("📤 Output")
    parsed_df = pd.DataFrame()
    source_label = ""

    if pasted and pasted.strip():
        parsed_df = parse_offer_text(pasted)
        source_label = supplier if supplier else "offerte"
    elif uploaded is not None:
        if uploaded.name.lower().endswith(".csv"):
            raw_df = pd.read_csv(uploaded)
        else:
            raw_df = pd.read_excel(uploaded)
        raw_df = raw_df.rename(columns={c: COLUMN_ALIASES.get(clean_text(c), clean_text(c)) for c in raw_df.columns})
        if len(raw_df.columns) == 1:
            parsed_df = parse_offer_text("\n".join(raw_df.iloc[:,0].astype(str).tolist()))
        else:
            parsed_df = parse_offer_text("\n".join(raw_df.astype(str).apply(" ".join, axis=1).tolist()))
        source_label = supplier_f if supplier_f else uploaded.name

    if not parsed_df.empty:
        total = len(parsed_df)
        ok = int((parsed_df["Review Flag"]=="NO").sum())
        review = total - ok
        m1, m2, m3 = st.columns(3)
        m1.metric("Totaal regels", total)
        m2.metric("OK", ok)
        m3.metric("Review nodig", review)
        st.dataframe(
            parsed_df[["Commodity","Product","GBX","Btls Case","Size CL","Cases MOQ","Purchase Price - Bottle","Purchase Price - Case","Currency","RF NRF","ST","Incoterms","Leadtime","Remark/BBD","Parse Status"]],
            use_container_width=True, height=380,
        )
        filename = f"JVH_{source_label.replace(' ','_')}_parsed.xlsx" if source_label else "JVH_parsed.xlsx"
        st.download_button(label="⬇️ Download Excel", data=to_excel_bytes(parsed_df), file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, type="primary")
        if review > 0:
            with st.expander(f"⚠️ {review} regels hebben review nodig"):
                st.dataframe(parsed_df[parsed_df["Review Flag"]=="YES"][["Product","Review Notes"]], use_container_width=True)
    else:
        st.info("Plak een offerte links of upload een bestand om te beginnen.")

st.markdown("---")
st.markdown('<p style="text-align:center; color:#555; font-size:0.8rem;">JVH Global B.V. · jvh-global.com</p>', unsafe_allow_html=True)
