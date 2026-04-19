"""JVH Global — Offer Intake Tool v2.0"""
from __future__ import annotations
import io, re
from decimal import Decimal, InvalidOperation
from typing import Any, Dict, List, Optional, Tuple
import pandas as pd
import streamlit as st

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------
st.set_page_config(page_title="JVH Offer Parser", page_icon="🥃", layout="wide")

JVH_COLUMNS = [
    "Commodity","Product","GBX","Btls Case","Size CL","ABV %","RF NRF","ST",
    "Cases MOQ","# btls case","Freight cost","Cost per case",
    "Purchase Price - Bottle","Purchase Price - Case","Margin case","Margin, %",
    "Price per bottle","Price per Case","Currency","Incoterms","Leadtime",
    "Remark/BBD","Source Row","Parse Status","Review Flag","Review Notes",
]

COMMODITY_MAP = {
    "absolut":"Vodka","belvedere":"Vodka","ciroc":"Vodka","grey goose":"Vodka","smirnoff":"Vodka",
    "rum":"Rum","bacardi":"Rum","captain morgan":"Rum","sailor jerry":"Rum",
    "gin":"Gin","tanqueray":"Gin","gin mare":"Gin","hendrick":"Gin","beefeater":"Gin","roku":"Gin",
    "bombay":"Gin","bombay sapphire":"Gin",
    "tequila":"Tequila","jose cuervo":"Tequila","1800":"Tequila","sierra":"Tequila",
    "don julio":"Tequila","olmeca":"Tequila","clase azul":"Tequila",
    "whisky":"Whisky","whiskey":"Whisky","jack daniel":"Whisky","jack daniels":"Whisky",
    "jim beam":"Whisky","jim beam white":"Whisky","jim beam cherry":"Whisky",
    "jim beam original":"Whisky","jim beam apple":"Whisky","jim beam honey":"Whisky",
    "teachers":"Whisky","famous grouse":"Whisky","glenfiddich":"Whisky","glenlivet":"Whisky",
    "hakushu":"Whisky","macallan":"Whisky","bowmore":"Whisky","dewar":"Whisky","dewars":"Whisky",
    "johnnie walker":"Whisky","hibiki":"Whisky","jameson":"Whisky","chivas":"Whisky",
    "grant's":"Whisky","grant s":"Whisky","lawson":"Whisky","auchentoshan":"Whisky","ballantines":"Whisky",
    "highland park":"Whisky","royal brackla":"Whisky","aultmore":"Whisky",
    "liqueur":"Liquor","liquor":"Liquor","aperol":"Liquor","jagermeister":"Liquor",
    "licor 43":"Liquor","kahlua":"Liquor","malibu":"Liquor",
    "cognac":"Cognac","hennessy":"Cognac","martell":"Cognac","camus":"Cognac",
    "remy martin":"Cognac","rémy martin":"Cognac",
    "champagne":"Champagne","veuve clicquot":"Champagne",
    "spritz":"RTD","wine":"Wines","sauvignon blanc":"Wines",
    "jacobs creek":"Wines","oyster bay":"Wines","brancott":"Wines",
    "mini":"Miniatures (5cl)","minis":"Miniatures (5cl)","miniatures":"Miniatures (5cl)",
}

SECTION_TO_COMMODITY = {
    "RUM":"Rum","TEQUILA":"Tequila","VODKA":"Vodka","WHISKY":"Whisky","WHISKEY":"Whisky",
    "GIN":"Gin","COGNAC":"Cognac","CHAMPAGNE":"Champagne","WINE":"Wines",
    "WINE (FCL)":"Wines","WINES":"Wines","LIQUOR":"Liquor","BEERS":"Beers","SOFTDRINKS":"Softdrinks",
}

COLUMN_ALIASES = {
    "Lead Time":"Leadtime","Warehouse":"Incoterms","Coded":"ST",
    "Cases Available (MOQ)":"Cases MOQ","Cases Available":"Cases MOQ",
    "RF/NRF":"RF NRF","REF/NRF":"RF NRF","producto":"Product","Producto":"Product",
    "btl/cs":"Btls Case","Btl/cs":"Btls Case","BTLS/CS":"Btls Case",
    "CL":"Size CL","alc %":"ABV %","ABV%":"ABV %",
    "Price":"Price per bottle","price":"Price per bottle",
    "cases":"Cases MOQ","CASES":"Cases MOQ","BRAND":"Product","SIZE LTR.":"Size LTR",
    "CAP":"RF NRF","STATUS":"ST","€/BTL":"Price per bottle",
    "EUROS/CASE":"Price per Case","ETA":"Leadtime",
}

CURRENCY_SYMBOLS = {"€":"EUR","$":"USD","eur":"EUR","usd":"USD","euro":"EUR"}

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def clean_text(v: Any) -> str:
    if v is None: return ""
    t = str(v).replace("\u00a0"," ").replace("–","-").replace("—","-").replace("\u2019","'")
    return re.sub(r"[ \t]+"," ",t).strip()

def parse_decimal(v: Any) -> Optional[Decimal]:
    t = clean_text(v).lower()
    if not t: return None
    for tok in ["eur","usd","euro","€","$","per bottle","per btl","per case","/btl","/cs"]:
        t = t.replace(tok,"")
    t = t.replace(" ","")
    if not t: return None
    if t.count(",") == 1 and t.count(".") >= 1: t = t.replace(".","").replace(",",".")
    elif t.count(",") == 1 and t.count(".") == 0: t = t.replace(",",".")
    try: return Decimal(t)
    except InvalidOperation: return None

def format_money(v: Optional[Decimal], currency: str) -> str:
    if v is None: return ""
    return f"{currency} {v:.2f}" if currency in ("EUR","USD") else f"{v:.2f}"

def detect_currency(*values: Any) -> str:
    joined = " ".join(clean_text(v).lower() for v in values if clean_text(v))
    for sym, code in CURRENCY_SYMBOLS.items():
        if sym in joined: return code
    return ""

def to_int(v: Any) -> Optional[int]:
    t = clean_text(v)
    if not t: return None
    m = re.search(r"\d+", t)
    return int(m.group()) if m else None

def to_float(v: Any) -> Optional[float]:
    d = parse_decimal(v)
    return float(d) if d is not None else None

def infer_commodity(product: str, size_cl: Optional[int] = None) -> str:
    p = clean_text(product).lower()
    if size_cl == 5 or "mini" in p or re.search(r"\b5\s?cl\b", p):
        return "Miniatures (5cl)"
    for key, commodity in COMMODITY_MAP.items():
        if key in p: return commodity
    return ""

def standardize_incoterms(v: str) -> str:
    t = clean_text(v)
    t = re.sub(r"(?i)\bexw\b","Exworks",t)
    t = re.sub(r"(?i)\bex\b","Exworks",t)
    return t

def standardize_leadtime(v: str) -> str:
    text = clean_text(v).replace("Lead time","").replace("Leadtime","").strip()
    t = text.lower().strip()
    if not t: return ""
    if re.search(r"(?i)\bon\s*floor\b|\bex\s*stock\b|\bin\s*stock\b|\bstock\b|\bready\b", t):
        return "On floor"
    m = re.match(r"(?i)^(mid|end|early|begin)\s+([a-z]+)$", t)
    if m: return m.group(1).capitalize()+" "+m.group(2).capitalize()
    m = re.match(r"(\d+)\s*-\s*(\d+)\s*days?", t)
    if m: return f"{m.group(1)}-{m.group(2)} Days"
    m = re.match(r"(\d+)\s*days?", t)
    if m: return f"{m.group(1)} Days"
    m = re.match(r"(\d+)\s*-\s*(\d+)\s*weeks?", t)
    if m:
        lo,hi = int(m.group(1)),int(m.group(2))
        return f"{lo*5}-{hi*5} Days"
    m = re.match(r"(\d+)\s*weeks?", t)
    if m:
        w = int(m.group(1))
        return f"{w*5}-{(w+1)*5} Days"
    return text

def standardize_rf(v: str) -> str:
    t = clean_text(v).upper().replace(".","")
    if not t: return "REF"
    if t in {"RF","REF","REFILLABLE"}: return "REF"
    if t in {"NRF","NON-REF","NON REF","NONREF"}: return "NRF"
    return t

def ensure_jvh_columns(df: pd.DataFrame) -> pd.DataFrame:
    for col in JVH_COLUMNS:
        if col not in df.columns: df[col] = ""
    return df[JVH_COLUMNS]

def build_output_row(data: Dict[str, Any]) -> Dict[str, Any]:
    row = {col:"" for col in JVH_COLUMNS}
    row.update(data)
    row["# btls case"] = row.get("Btls Case","")
    row["Price per bottle"] = row.get("Purchase Price - Bottle","")
    row["Price per Case"] = row.get("Purchase Price - Case","")
    return row

# ---------------------------------------------------------------------------
# Blob splitter — voor samengesmolten PDF-kopieën
# ---------------------------------------------------------------------------
def detect_and_split_blob(text: str) -> str:
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    blob_lines = [l for l in lines if len(l) > 150 and l.count("€") >= 2]
    if not blob_lines:
        return text
    result_lines = [l for l in lines if l not in blob_lines]
    for blob in blob_lines:
        blob = re.sub(r"(?i)Description\s*QTY\s*BOTT[A-Z\s]*Lead\s*time\s*","",blob)
        # Split eerst op scheiders
        parts = re.split(r"(?:(?:after\s+)?deposit|LOEND\.?)(?=[A-Z])", blob)
        for part in parts:
            part = part.strip().rstrip(".")
            if not part or len(part) < 5: continue
            part = re.sub(r"(\d+(?:\.\d+)?)[Ll]\s*[xX]\s*(\d+)", lambda m: f"{m.group(2)}x{m.group(1)}L", part)
            part = re.sub(r"(\d+)[Mm][Ll]\s*[xX*]\s*(\d+)", lambda m: f"{m.group(2)}x{m.group(1)}ml", part)
            part = re.sub(r"(\d+)[Ll]\s*\*\s*(\d+)", lambda m: f"{m.group(2)}x{m.group(1)}L", part)
            part = re.sub(r"(?i)\bON THE FLOOR\s*LOEND\.?","On floor Exworks Loendersloot",part)
            part = re.sub(r"(?i)\bON THE FLOOR\b","On floor",part)
            part = re.sub(r"(?i)\b1-2 weeks after\b","1-2 weeks",part)
            part = re.sub(r"€\s*([\d,]+(?:\.\d+)?)", lambda m: f"@ EUR {m.group(1).replace(',','')} /btl", part)
            part = re.sub(r"\$\s*([\d,]+(?:\.\d+)?)", lambda m: f"@ USD {m.group(1).replace(',','')} /btl", part)
            if not re.search(r"(?i)exworks|DAP|CFR|EXW", part): part += " Exworks Loendersloot"
            if "CODED" not in part.upper(): part += " CODED"
            result_lines.append(part)
    return "\n".join(result_lines)

# ---------------------------------------------------------------------------
# Preprocessor
# ---------------------------------------------------------------------------
def preprocess_text(text: str) -> str:
    for old, new in {
        "RF.":"RF","/cs.":"/cs","/btl.":"/btl","/btl.,":"/btl,","/cs.,":"/cs,",
        " at ":" @ "," per bottle":" /btl"," per case":" /cs",
        " per cs":" /cs"," per btl":" /btl",
    }.items():
        text = text.replace(old, new)
    cleaned = []
    for raw in text.splitlines():
        line = raw.strip()
        if not line:
            cleaned.append("")
            continue
        line = re.sub(r"^[-*•]\s*","",line)
        line = re.sub(r"(?<=\d),(?=\d{3}\b)","",line)
        line = line.replace("–","-").replace("—","-")
        line = re.sub(r"\((\d+(?:\.\d+)?)[Ll]\s*[xX]\s*(\d+)\)", lambda m: f"{m.group(2)}x{m.group(1)}L", line)
        line = re.sub(r"\((\d+)[Cc][Ll]\s*[xX]\s*(\d+)\)", lambda m: f"{m.group(2)}x{m.group(1)}cl", line)
        line = re.sub(r"(?i)\bQty\s*:\s*","",line)
        line = re.sub(r"(?i)\bPrice\s*:\s*€\s*(\d+(?:[.,]\d+)?)\s*/(btl|cs)\b",r"@ EUR \1 /\2",line)
        line = re.sub(r"(?i)\bPrice\s*:\s*\$\s*(\d+(?:[.,]\d+)?)\s*/(btl|cs)\b",r"@ USD \1 /\2",line)
        line = re.sub(r"(?i)\bPrice\s*:\s*(USD|EUR)\s*(\d+(?:[.,]\d+)?)\s*/(btl|cs)\b",r"@ \1 \2 /\3",line)
        line = re.sub(r"€\s*(\d+(?:[.,]\d+)?)\s*/(btl|cs)\b",r"@ EUR \1 /\2",line)
        line = re.sub(r"\$\s*(\d+(?:[.,]\d+)?)\s*/(btl|cs)\b",r"@ USD \1 /\2",line)
        line = re.sub(r"(?i)\bex-([A-Za-z]+)",r"ex \1",line)
        line = re.sub(r"(?i)\bDuty\s*Status\s*:\s*","",line)
        line = re.sub(r"(?i)\beuro\s+(\d)",r"@ EUR \1",line)
        line = re.sub(r"(?i)\beuros\s+(\d)",r"@ EUR \1",line)
        line = re.sub(r"(?i)-\s*(\d+(?:[.,]\d+)?)\s*(USD|EUR|€|\$)\b",r"@ \2 \1",line)
        line = re.sub(r"(?i)\b(\d+(?:[.,]\d+)?)\s*(USD|EUR)\s*(?:per case|/cs)?\s*$",r"@ \2 \1 /cs",line)
        line = re.sub(r"(?i)\b(USD|EUR)\s+(\d+(?:[.,]\d+)?)\s+per case\b",r"@ \1 \2 /cs",line)
        line = re.sub(r"(?i)\b(USD|EUR)\s+(\d+(?:[.,]\d+)?)\s*/(cs|btl)\b",r"@ \1 \2 /\3",line)
        if re.match(r"(?i)^FTL\s+",line) and not re.search(r"(?i)\d+\s*(cases|case|cs|bottles|bottle|btls)\b",line):
            line = line + " 1 cs FTL_LINE"
        cleaned.append(line)
    return "\n".join(cleaned)

# ---------------------------------------------------------------------------
# Main parser
# ---------------------------------------------------------------------------
def parse_offer_text(text: str) -> pd.DataFrame:
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
            current_section = SECTION_TO_COMMODITY[upper.rstrip(":")]
            continue
        if upper.startswith("BBD:") or "MOQ:" in upper:
            m = re.search(r"(?i)bbd:\s*([^-]+(?:[-/][^-]+)?)", line)
            if m: trailing_bbd = clean_text(m.group(1))
            m = re.search(r"(?i)moq:\s*([\d,]+)\s*cs", line)
            if m: trailing_moq = int(m.group(1).replace(",",""))
            continue
        if "CODED" in upper and "ALL ITEMS" in upper:
            default_coded = True
            continue

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
            btls_case = int(m.group(1))
            size_cl = int(Decimal(m.group(2)) * Decimal("100"))
        else:
            m = re.search(r"(?i)(\d+)x(\d+(?:\.\d+)?)ml\b", line)
            if m:
                btls_case = int(m.group(1))
                size_cl = int(Decimal(m.group(2)) / Decimal("10"))
            else:
                m = re.search(r"(?i)(\d+)x(\d+)cl\b", line)
                if m:
                    btls_case = int(m.group(1))
                    size_cl = int(m.group(2))

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
            lead_match = re.search(r"(?i)(on\s*floor|stock|ready|\d+\s*-\s*\d+\s*(?:weeks?|days?)|\d+\s*(?:weeks?|days?)|mid\s+[a-z]+|end\s+[a-z]+|early\s+[a-z]+)", line)
            if lead_match: leadtime = standardize_leadtime(lead_match.group(1))
        if not leadtime and global_leadtime: leadtime = global_leadtime

        product = line
        for p in [
            r"(?i)\bFTL\b", r"\bFTL_LINE\b", r"\b1\s+cs\b",
            r"(?i)\b[\d,]+\s*(cases|case|cs|bottles|bottle|btls)\b",
            r"(?i)@\s*(EUR|USD|€|\$)?\s*[0-9]+(?:[.,][0-9]+)?\s*(?:/(?:btl|cs)|per\s+(?:bottle|case))?",
            r"(?i)\b(EX(?:W| |\s)\s*[A-Za-z]+(?:\s+[A-Za-z]+)?|DAP\s+[A-Za-z]+(?:\s+[A-Za-z]+)?|CFR\s+[A-Za-z]+|CNF\s+[A-Za-z]+)\b",
            r"(?i)\b(on\s*floor|stock|ready|\d+\s*-\s*\d+\s*(?:weeks?|days?)|\d+\s*(?:weeks?|days?)|mid\s+[a-z]+)\b",
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

def to_excel_bytes(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, sheet_name="JVH Master", index=False)
        df[df["Review Flag"]=="YES"].copy().to_excel(w, sheet_name="Review", index=False)
    buf.seek(0)
    return buf.getvalue()

# ---------------------------------------------------------------------------
# UI
# ---------------------------------------------------------------------------
st.title("🥃 JVH Global — Offer Parser")
st.caption("Plak een offerte of upload een bestand → download de Excel")

col_in, col_out = st.columns([1, 1])

with col_in:
    st.subheader("📥 Input")
    tab_text, tab_file = st.tabs(["Tekst plakken", "Bestand uploaden"])

    with tab_text:
        pasted = st.text_area(
            "Plak hier de offertetekst (email, WhatsApp, etc.)",
            height=400,
            placeholder="Plak de volledige offertetekst hier...",
        )
        supplier = st.text_input("Leverancier (optioneel)", placeholder="bijv. Diageo / Pernod / ...")

    with tab_file:
        uploaded = st.file_uploader("Upload Excel of CSV", type=["xlsx","xls","csv"])
        supplier_f = st.text_input("Leverancier (optioneel) ", placeholder="bijv. Diageo / Pernod / ...")

with col_out:
    st.subheader("📤 Output")

    parsed_df = pd.DataFrame()
    source_label = ""

    # Parse from text
    if pasted and pasted.strip():
        parsed_df = parse_offer_text(pasted)
        source_label = supplier if supplier else "offerte"

    # Parse from file
    elif uploaded is not None:
        if uploaded.name.lower().endswith(".csv"):
            raw_df = pd.read_csv(uploaded)
        else:
            raw_df = pd.read_excel(uploaded)
        raw_df = raw_df.rename(columns={c: COLUMN_ALIASES.get(clean_text(c), clean_text(c)) for c in raw_df.columns})
        # Try as free text first if single column
        if len(raw_df.columns) == 1:
            text_blob = "\n".join(raw_df.iloc[:,0].astype(str).tolist())
            parsed_df = parse_offer_text(text_blob)
        else:
            from typing import Tuple
            def parse_table_dataframe(df):
                cols = {clean_text(c) for c in df.columns}
                if {"Product","Btls Case","Size LTR","ABV %","RF NRF","ST","Price per bottle","Price per Case","Leadtime","Cases MOQ"}.issubset(cols):
                    return df, "inventory_eta_table"
                return df, "unknown"
            parsed_df = parse_offer_text("\n".join(raw_df.astype(str).apply(" ".join, axis=1).tolist()))
        source_label = supplier_f if supplier_f else uploaded.name

    if not parsed_df.empty:
        total = len(parsed_df)
        ok = int((parsed_df["Review Flag"]=="NO").sum())
        review = total - ok

        m1, m2, m3 = st.columns(3)
        m1.metric("Totaal regels", total)
        m2.metric("✅ OK", ok)
        m3.metric("⚠️ Review", review)

        st.dataframe(
            parsed_df[["Commodity","Product","GBX","Btls Case","Size CL","Cases MOQ",
                       "Purchase Price - Bottle","Purchase Price - Case","Currency",
                       "RF NRF","ST","Incoterms","Leadtime","Remark/BBD","Parse Status"]],
            use_container_width=True,
            height=380,
        )

        filename = f"JVH_{source_label.replace(' ','_')}_parsed.xlsx" if source_label else "JVH_parsed.xlsx"
        st.download_button(
            label="⬇️ Download Excel",
            data=to_excel_bytes(parsed_df),
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary",
        )

        if review > 0:
            with st.expander(f"⚠️ {review} regels hebben review nodig"):
                st.dataframe(
                    parsed_df[parsed_df["Review Flag"]=="YES"][["Product","Review Notes"]],
                    use_container_width=True,
                )
    else:
        st.info("Plak een offerte links of upload een bestand om te beginnen.")
