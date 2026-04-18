import re
import zipfile
import tempfile
from datetime import datetime, date
from io import BytesIO
from pathlib import Path

import pandas as pd
import pdfplumber
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Next Supply | Separador de OPS", page_icon="N", layout="wide")

st.markdown("""<style>
:root{
    --bg:#02053d;--panel:#060a53;--panel2:#08106a;
    --border:rgba(63,93,255,0.35);--border-soft:rgba(255,255,255,0.08);
    --text:#f5f7ff;--muted:#a9b4ea;
    --primary:#1237ff;--primary2:#3f5dff;
    --shadow:0 18px 60px rgba(0,0,0,0.28);
    --radius-xl:28px;--radius-lg:22px;
}
.stApp{
    background:
        radial-gradient(circle at top left,rgba(63,93,255,0.20),transparent 28%),
        radial-gradient(circle at top right,rgba(18,55,255,0.16),transparent 24%),
        linear-gradient(180deg,#010332 0%,var(--bg) 100%);
    color:var(--text);
}
.block-container{max-width:1400px;padding-top:1.4rem;padding-bottom:2.5rem;}
[data-testid="stHeader"]{background:transparent;}
[data-testid="stSidebar"]{background:var(--panel) !important;border-right:1px solid var(--border-soft) !important;}
[data-testid="stSidebar"] *{color:var(--text) !important;}
h1,h2,h3,h4,h5,h6,p,label,div,span{color:var(--text);}
.topbar{display:flex;align-items:center;justify-content:space-between;gap:24px;background:rgba(4,8,70,0.80);border:1px solid var(--border-soft);border-radius:24px;padding:18px 24px;box-shadow:var(--shadow);margin-bottom:22px;}
.brand-wrap{display:flex;align-items:center;gap:16px;}
.brand-mark{width:58px;height:58px;border-radius:14px;background:linear-gradient(135deg,var(--primary2),var(--primary));display:flex;align-items:center;justify-content:center;color:white !important;font-size:1.8rem;font-weight:800;box-shadow:0 10px 30px rgba(18,55,255,0.35);}
.brand-title{font-size:1.5rem;font-weight:800;line-height:1;margin:0;}
.brand-subtitle{margin:4px 0 0 0;color:var(--muted) !important;font-size:0.95rem;}
.hero{background:linear-gradient(180deg,rgba(6,10,83,0.92) 0%,rgba(4,7,58,0.92) 100%);border:1px solid var(--border);border-radius:var(--radius-xl);padding:30px;box-shadow:var(--shadow);margin-bottom:20px;}
.hero-title{font-size:2.2rem;font-weight:800;margin-bottom:0.4rem;}
.hero-sub{color:var(--muted) !important;font-size:1rem;margin-bottom:1.4rem;max-width:900px;}
.metrics{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:16px;}
.mcard{background:rgba(255,255,255,0.04);border:1px solid var(--border-soft);border-radius:var(--radius-lg);padding:18px 20px;backdrop-filter:blur(8px);}
.mcard-label{color:var(--muted) !important;font-size:0.9rem;margin-bottom:8px;}
.mcard-val{font-size:1.7rem;font-weight:800;line-height:1.1;color:var(--text) !important;}
.mcard-val.accent{color:#7b9fff !important;}
.sec-title{font-size:1rem;font-weight:700;border-left:3px solid var(--primary2);padding-left:12px;margin:20px 0 12px 0;}
.alerta-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:14px;margin-bottom:4px;}
.alerta-card{border-radius:16px;padding:20px 22px;display:flex;flex-direction:column;gap:6px;}
.alerta-card.red{background:rgba(220,38,38,0.15);border:1px solid rgba(220,38,38,0.45);}
.alerta-card.yellow{background:rgba(234,179,8,0.12);border:1px solid rgba(234,179,8,0.40);}
.alerta-card.green{background:rgba(34,197,94,0.10);border:1px solid rgba(34,197,94,0.35);}
.alerta-label{font-size:0.8rem;font-weight:600;opacity:0.75;}
.alerta-val{font-size:2rem;font-weight:900;line-height:1;}
.alerta-card.red .alerta-val{color:#f87171 !important;}
.alerta-card.yellow .alerta-val{color:#fbbf24 !important;}
.alerta-card.green .alerta-val{color:#4ade80 !important;}
.alerta-sub{font-size:0.72rem;opacity:0.6;}
div[data-testid="stFileUploader"] section{background:rgba(255,255,255,0.03) !important;border:1px dashed rgba(88,117,255,0.45) !important;border-radius:18px !important;padding:8px !important;}
div[data-testid="stDataFrame"]{border:1px solid var(--border-soft) !important;border-radius:18px !important;overflow:hidden !important;background:rgba(255,255,255,0.02) !important;}
.stDownloadButton>button,.stButton>button{border-radius:14px !important;min-height:46px !important;padding:0.72rem 1.15rem !important;font-weight:800 !important;border:1px solid transparent !important;color:white !important;background:linear-gradient(135deg,var(--primary),var(--primary2)) !important;box-shadow:0 10px 24px rgba(18,55,255,0.28) !important;}
.stAlert{border-radius:16px !important;}
.hr{height:1px;background:linear-gradient(90deg,var(--primary2),rgba(63,93,255,0.1),transparent);border:none;margin:24px 0;}
@media(max-width:1100px){.metrics{grid-template-columns:repeat(2,minmax(0,1fr));}.alerta-grid{grid-template-columns:1fr;}}
@media(max-width:700px){.metrics{grid-template-columns:1fr;}.hero-title{font-size:1.8rem;}}
</style>""", unsafe_allow_html=True)


# ── Utilitarios ───────────────────────────────────────────────────────────────

def clean(text):
    text = (text or "").replace("\u00ad", "")
    text = re.sub(r"Resumo extra[i\u00ed]do por[^\n]*(NEXTSUPPLY\d+\))[^\n]*", "", text, flags=re.I)
    text = re.sub(r"\bP[a\u00e1]g[.:]\s*\d+/\d+\b", "", text, flags=re.I)
    text = re.sub(r"(?is)Resumo extra[i\u00ed]do por.*?(?=\n|$)", "", text)
    # Remove cabecalho de pagina do Petronect que aparece no meio dos blocos
    text = re.sub(
        r"(?im)^Resumo da Oportunidade\s+N[u\u00fa]mero da Oportunidade\s*\n.*?\d{10}\s*\n?",
        "",
        text,
    )
    text = re.sub(
        r"(?im)^N[u\u00fa]mero da Oportunidade\s*\n\s*\d{10}\s*\n\s*Resumo da Oportunidade\s*\n[^\n]+\n?",
        "",
        text,
    )
    text = re.sub(r"(?im)^\s*Resumo da Oportunidade\s*$", "", text)
    text = re.sub(r"(?is)ENDERE[C\u00c7]O DE ENTREGA E FATURAMENTO:.*?(?=\nDados do Item|\nDescri|\nDeclara|\Z)", "", text)
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\s*\n\s*", "\n", text)
    return text.strip()


def one_line(text):
    text = clean(text).replace("\n", " ")
    text = re.sub(r"\(NEXTSUPPLY\d+\)", "", text, flags=re.I)
    return re.sub(r"\s{2,}", " ", text).strip()


def brand_regex(term):
    parts = re.split(r"\s+", str(term).strip().lower())
    return r"\b" + r"[\s\-]?".join(map(re.escape, parts)) + r"\b"


# ── Categorias ────────────────────────────────────────────────────────────────

CATEGORIAS = [
    ("Valvula",           ["valv", "ball valve", "gate valve", "check valve", "globo", "borboleta", "esfera", "gaveta", "retencao", "alivio"]),
    ("Tubo",              ["tubo ", "tubulac", " piping", "conduc", " pipe "]),
    ("Junta / Vedacao",   ["junta", "gaxeta", "oring", "o-ring", "gasket", "vedac", "anel de veda", "anel lip"]),
    ("Rele",              ["rele ", "relay", "rele de protecao", "rele de tensao", "rele aux"]),
    ("Bomba",             ["bomba", " pump ", "booster"]),
    ("Filtro",            ["filtro", "filter", "coador", "strainer"]),
    ("Instrumento",       ["instrumento", "sensor", "transmissor", "medidor", "manometro", "pressostato", "termometro"]),
    ("Cabo / Fio",        ["cabo ", "cable", "condutor", "teldor"]),
    ("Flange",            ["flange"]),
    ("Conector / Fitting",["fitting", "niple", "cotovelo", " tee "]),
    ("Parafuso / Fixador",["parafuso", "arruela", " bolt", " stud "]),
    ("Mangueira",         ["mangueira", " hose "]),
    ("Conjunto / Kit",    ["conj.", "conjunto especif", " kit ", "assembly"]),
    ("Relatorio / Doc",   ["manual de operac", "manual tecnico", "documento tecnico", "relatorio tecnico"]),
]


def categorize(text):
    t = (text or "").lower()
    for cat, keywords in CATEGORIAS:
        for kw in keywords:
            if kw in t:
                return cat
    return "Outros"


# ── Deteccao de recorrencia ───────────────────────────────────────────────────

def detect_recurring(df_new, df_hist):
    if df_hist.empty or df_new.empty:
        return pd.Series([False] * len(df_new), index=df_new.index)

    def norm(s):
        return re.sub(r"\s+", " ", str(s or "").lower().strip())

    hist_fab  = set(df_hist["Fabricante/PN"].dropna().map(norm)) - {"", "nan"}
    hist_desc = set(df_hist["Descricao longa do item"].dropna().map(norm)) - {"", "nan"}

    def is_rec(row):
        fab  = norm(row.get("Fabricante/PN", ""))
        desc = norm(row.get("Descricao longa do item", ""))
        return (fab and fab in hist_fab) or (desc and desc in hist_desc)

    return df_new.apply(is_rec, axis=1)


# ── Alertas de prazo ─────────────────────────────────────────────────────────

def classify_prazo(data_str):
    try:
        d = datetime.strptime(str(data_str).strip(), "%d/%m/%Y").date()
        delta = (d - date.today()).days
        if delta <= 1:   return delta, "red"
        elif delta <= 3: return delta, "yellow"
        else:            return delta, "green"
    except Exception:
        return None, "green"


def render_alertas(df):
    if df is None or df.empty or "Data (cotacao)" not in df.columns:
        return
    ops = df.drop_duplicates(subset=["Numero da Oportunidade"])[["Numero da Oportunidade", "Data (cotacao)"]].copy()
    ops["_dias"], ops["_cor"] = zip(*ops["Data (cotacao)"].map(classify_prazo))
    n_red    = int((ops["_cor"] == "red").sum())
    n_yellow = int((ops["_cor"] == "yellow").sum())
    n_green  = int((ops["_cor"] == "green").sum())

    sec("Alertas de Prazo")
    st.markdown(
        '<div class="alerta-grid">'
        '<div class="alerta-card red"><div class="alerta-label">URGENTE - Vence hoje / amanha</div>'
        '<div class="alerta-val">' + str(n_red) + '</div>'
        '<div class="alerta-sub">' + (str(n_red) + " oportunidade(s)" if n_red else "Nenhuma urgente") + '</div></div>'
        '<div class="alerta-card yellow"><div class="alerta-label">ATENCAO - Vence em 2 a 3 dias</div>'
        '<div class="alerta-val">' + str(n_yellow) + '</div>'
        '<div class="alerta-sub">' + (str(n_yellow) + " oportunidade(s)" if n_yellow else "Nenhuma em atencao") + '</div></div>'
        '<div class="alerta-card green"><div class="alerta-label">OK - Prazo confortavel</div>'
        '<div class="alerta-val">' + str(n_green) + '</div>'
        '<div class="alerta-sub">' + (str(n_green) + " oportunidade(s)" if n_green else "Nenhuma") + '</div></div>'
        '</div>',
        unsafe_allow_html=True
    )
    if n_red > 0:
        red_nums = ops[ops["_cor"] == "red"]["Numero da Oportunidade"]
        with st.expander("Ver oportunidades urgentes"):
            urgentes = df[df["Numero da Oportunidade"].isin(red_nums)]
            show_cols = [c for c in ["Numero da Oportunidade", "Data (cotacao)", "Responsavel", "Descricao longa do item"] if c in urgentes.columns]
            st.dataframe(urgentes[show_cols], use_container_width=True, hide_index=True)


# ── Extracao PDF ──────────────────────────────────────────────────────────────

def extract_header(raw, fallback=""):
    m = re.search(r"N[u\u00fa]mero da Oportunidade\s*(\d{10})", raw, flags=re.S | re.I)
    if m:
        numero = m.group(1)
    else:
        fb = re.search(r"(\d{10})", str(fallback))
        numero = fb.group(1) if fb else ""

    def get(pattern):
        r = re.search(pattern, raw, flags=re.I)
        return r.group(1).strip() if r else ""

    tipo = get(r"Tipo de Oportunidade\s*([^\n]+)")
    crit = get(r"Crit[e\u00e9]rio de Julgamento\s*([^\n]+)")
    fim_raw = get(r"Fim do per[i\u00ed]odo de cota[c\u00e7][a\u00e3]o\s*([0-9]{2}\.[0-9]{2}\.[0-9]{4}\s*/\s*[0-9]{2}:[0-9]{2}:[0-9]{2})")
    fim = re.sub(r"\s*/\s*", " / ", fim_raw)
    ml = re.search(r"Local de Entrega\s*(.*?)\nInforma[c\u00e7][o\u00f5]es do Comprador", raw, flags=re.S | re.I)
    local = "; ".join(x.strip() for x in ml.group(1).split("\n") if x.strip()) if ml else ""
    return {"numero": numero, "tipo": tipo, "crit": crit, "fim": fim, "local": local}


def extract_item_id(block):
    for pat, flags in [
        (r"(?m)^\s*(\d{1,6})\s+\S", 0),
        (r"(?i)N[u\u00fa]mero\s+Descri.*?\n(\d+)\s", re.S),
        (r"(?i)N[u\u00fa]mero\s*\n\s*do item\s*(?:\n|\s)+([0-9A-Za-z\.\-]+)", 0),
    ]:
        m = re.search(pat, block, flags=flags)
        if m:
            return re.sub(r"\D", "", m.group(1))
    return ""


def extract_qty_unit(block, item_id):
    if not item_id:
        return "", ""
    patterns = [
        rf"(?im)^\s*{re.escape(item_id)}\s+.*?\bMaterial\b\s+([0-9]+(?:\.[0-9]{{3}})*(?:,[0-9]+)?)\s+([A-Za-z\u00c0-\u00ff]+)\s+\d{{2}}\.\d{{2}}\.\d{{4}}\b",
        rf"(?im)^\s*{re.escape(item_id)}\s+.*?\b([0-9]+(?:\.[0-9]{{3}})*(?:,[0-9]+)?)\s+([A-Za-z\u00c0-\u00ff]+)\s+\d{{2}}\.\d{{2}}\.\d{{4}}\b",
        rf"(?is)\b{re.escape(item_id)}\b.*?\bMaterial\b\s+([0-9]+(?:\.[0-9]{{3}})*(?:,[0-9]+)?)\s+([A-Za-z\u00c0-\u00ff]+)\b",
        r"(?is)\bQuantidade\b\s*([0-9]+(?:\.[0-9]{3})*(?:,[0-9]+)?)\s+\b([A-Za-z\u00c0-\u00ff]+)\b",
    ]
    for p in patterns:
        m = re.search(p, block)
        if m:
            qty = m.group(1).strip()
            qty_clean = qty.replace(".", "").replace(",", ".")
            try:
                qty_num = float(qty_clean)
                qty_fmt = str(int(qty_num)) if qty_num == int(qty_num) else str(qty_num)
            except Exception:
                qty_fmt = qty
            return qty_fmt, m.group(2).strip()
    return "", ""


def extract_desc(block):
    b = re.split(r"(?i)ENDERE[C\u00c7]O DE ENTREGA", block, maxsplit=1)[0]
    m = re.search(
        r"(?is)Descri[c\u00e7][a\u00e3]o de Item\s*(.*?)(?:\nDescri[c\u00e7][a\u00e3]o longa|\Z)",
        b
    )
    result = one_line(m.group(1)) if m else ""
    return result if result.strip() else "Sem descricao"


def extract_long(block):
    b = re.split(r"(?i)ENDERE[C\u00c7]O DE ENTREGA", block, maxsplit=1)[0]
    b = re.split(r"(?i)Declara[c\u00e7][o\u00f5]es envolvidas", b, maxsplit=1)[0]
    b = re.split(r"(?i)Resumo extra[i\u00ed]do por", b, maxsplit=1)[0]
    b = re.split(r"(?i)Resumo da Oportunidade", b, maxsplit=1)[0]
    m = re.search(
        r"(?is)Descri[c\u00e7][a\u00e3]o longa do item\s*(.*?)$|Descri[c\u00e7][a\u00e3]o longa\s*(.*?)$",
        b
    )
    if not m:
        return ""
    result = one_line(m.group(1) or m.group(2))
    result = re.sub(r"\s*\bP[a\u00e1]g:\s*\d+/\d+.*$", "", result, flags=re.I).strip()
    result = re.sub(r"\s*\b\d{2}\.\d{2}\.\d{4}\s*(?:[a\u00e0]s)?\s*\d{2}:\d{2}:\d{2}.*$", "", result, flags=re.I).strip()
    result = re.sub(r"\s*\(NEXTSUPPLY\d+\).*$", "", result, flags=re.I).strip()
    result = re.sub(r"[\s;,]+$", "", result).strip()
    return result


def extract_manufacturer(block):
    b = re.split(r"(?i)ENDERE[C\u00c7]O DE ENTREGA", block, maxsplit=1)[0]
    b = re.split(r"(?i)Resumo extra[i\u00ed]do por", b, maxsplit=1)[0]
    b = re.split(r"(?i)Resumo da Oportunidade", b, maxsplit=1)[0]
    m = re.search(r"(?i)Tp:\s*(.+?)(?=\s*-{5,}|\s*ENDERE|\s*Resumo|\s*Dados|\s*Declara|$)", b, re.S)
    if m:
        fab = re.sub(r"\s{2,}", " ", m.group(1)).strip()
        fab = re.sub(r"[\s/\-|:;,]+$", "", fab).strip()
        if fab:
            return fab
    m2 = re.search(r"(?i)FABRICANTE:\s*(\S+)", b)
    if m2:
        return m2.group(1).strip()
    m3 = re.search(r"(?i)REFER[E\u00ca]NCIA:\s*([^\n/]+)", b)
    if m3:
        return m3.group(1).strip()
    return ""


# ── Atribuicao (logica original NextSupply) ───────────────────────────────────

HELIO_BRANDS  = ["abb", "schneider", "siemens", "rittal", "phoenix", "weidmuller", "rockwell"]
MAYARA_BRANDS = ["skf", "emerson"]
VIVIANA_BRANDS= ["kongsberg", "yamada", "dnh", "evac", "steyr"]

RESPONSAVEIS  = ["Viviana", "Helio", "Mayara"]
TARGETS_PCT   = {"Viviana": 0.33, "Helio": 0.33, "Mayara": 0.34}


def assign(df):
    if df.empty:
        df["Responsavel"] = pd.Series(dtype="object")
        return df

    corpus = (
        df["Descricao de Item"].fillna("") + " " +
        df["Descricao longa do item"].fillna("") + " " +
        df["Fabricante/PN"].fillna("")
    ).str.lower()

    df["Responsavel"] = pd.NA

    # Viviana - tipo Inaplicavel
    df.loc[
        df["Tipo de Oportunidade"].fillna("").str.contains("Inaplica", na=False, regex=False),
        "Responsavel",
    ] = "Viviana"

    # Helio - marcas eletrica/automacao
    for t in HELIO_BRANDS:
        mask = df["Responsavel"].isna() & corpus.str.contains(brand_regex(t), regex=True, na=False)
        df.loc[mask, "Responsavel"] = "Helio"

    # Mayara - marcas especificas
    for t in MAYARA_BRANDS:
        mask = df["Responsavel"].isna() & corpus.str.contains(brand_regex(t), regex=True, na=False)
        df.loc[mask, "Responsavel"] = "Mayara"

    # Viviana - marcas nauticas
    for t in VIVIANA_BRANDS:
        mask = df["Responsavel"].isna() & corpus.str.contains(brand_regex(t), regex=True, na=False)
        df.loc[mask, "Responsavel"] = "Viviana"

    # Balanceamento proporcional para os restantes
    total = len(df)
    targets = {r: round(total * TARGETS_PCT[r]) for r in RESPONSAVEIS}
    targets["Mayara"] = total - targets["Viviana"] - targets["Helio"]
    counts = df["Responsavel"].value_counts(dropna=True).to_dict()
    for r in RESPONSAVEIS:
        counts.setdefault(r, 0)

    chave = (
        df["Descricao de Item"].fillna("").str.strip().str.lower() + "|" +
        df["Descricao longa do item"].fillna("").str.strip().str.lower() + "|" +
        df["Fabricante/PN"].fillna("").str.strip().str.lower()
    )

    for _, grp in df[df["Responsavel"].isna()].groupby(chave, sort=False):
        chosen = max(RESPONSAVEIS, key=lambda r: targets[r] - counts[r])
        df.loc[grp.index, "Responsavel"] = chosen
        counts[chosen] += len(grp)

    return df


# ── Pipeline ──────────────────────────────────────────────────────────────────

COLS = [
    "Numero da Oportunidade", "Tipo de Oportunidade", "Criterio de Julgamento",
    "Fim do periodo de cotacao", "Data (cotacao)", "Hora (cotacao)",
    "Local de Entrega", "Item", "Quantidade", "Unidade de medida",
    "Descricao de Item", "Descricao longa do item",
    "Fabricante/PN", "Categoria", "Recorrente", "Responsavel",
]


def process_zip(zip_bytes, df_hist=None):
    rows, log = [], []
    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        with zipfile.ZipFile(BytesIO(zip_bytes)) as z:
            z.extractall(tmp)
        pdfs = sorted(tmp.rglob("*.pdf"))
        if not pdfs:
            return pd.DataFrame(columns=COLS), ["Nenhum PDF encontrado no ZIP."]
        for pdf in pdfs:
            try:
                with pdfplumber.open(str(pdf)) as p:
                    raw = clean("\n".join((pg.extract_text() or "") for pg in p.pages))
                header = extract_header(raw, pdf.stem)
                blocks = re.split(r"(?i)\bDados do Item\b", raw)[1:]
                if not blocks:
                    log.append("AVISO: " + pdf.name + " sem itens.")
                    continue
                count = 0
                for block in blocks:
                    item_id = extract_item_id(block)
                    if not item_id:
                        continue
                    qty, unit = extract_qty_unit(block, item_id)
                    desc   = extract_desc(block)
                    long_d = extract_long(block)
                    fab    = extract_manufacturer(block)
                    rows.append({
                        "Numero da Oportunidade":    header["numero"],
                        "Tipo de Oportunidade":      header["tipo"],
                        "Criterio de Julgamento":    header["crit"],
                        "Fim do periodo de cotacao": header["fim"],
                        "Local de Entrega":          header["local"],
                        "Item":                      item_id,
                        "Quantidade":                qty,
                        "Unidade de medida":         unit,
                        "Descricao de Item":         desc,
                        "Descricao longa do item":   long_d,
                        "Fabricante/PN":             fab,
                        "Categoria":                 categorize(long_d or desc),
                    })
                    count += 1
                log.append("OK: " + pdf.name + " - " + str(count) + " item(ns).")
            except Exception as e:
                log.append("ERRO: " + pdf.name + " - " + str(e))

    if not rows:
        return pd.DataFrame(columns=COLS), log

    df = pd.DataFrame(rows).drop_duplicates(subset=["Numero da Oportunidade", "Item"], keep="first")
    dt = pd.to_datetime(
        df["Fim do periodo de cotacao"].astype(str).str.replace(" / ", " ", regex=False),
        format="%d.%m.%Y %H:%M:%S", errors="coerce",
    )
    df["Data (cotacao)"] = dt.dt.strftime("%d/%m/%Y")
    df["Hora (cotacao)"] = dt.dt.strftime("%H:%M:%S")

    hist = df_hist if df_hist is not None and not df_hist.empty else pd.DataFrame()
    df["Recorrente"] = detect_recurring(df, hist).map({True: "Sim", False: "Nao"})

    df = assign(df)
    for col in COLS:
        if col not in df.columns:
            df[col] = ""
    return df[COLS], log


# ── Excel formatado ───────────────────────────────────────────────────────────

HDR_FILL  = PatternFill("solid", fgColor="02053D")
HDR_FONT  = Font(name="Arial", bold=True, color="FFFFFF", size=10)
ODD_FILL  = PatternFill("solid", fgColor="F0F2FF")
EVEN_FILL = PatternFill("solid", fgColor="FFFFFF")
REC_FILL  = PatternFill("solid", fgColor="FFF3CD")
BODY_FONT = Font(name="Arial", size=10)
CENTER    = Alignment(horizontal="center", vertical="center", wrap_text=False)
LEFT      = Alignment(horizontal="left", vertical="center", wrap_text=True)
THIN      = Side(style="thin", color="C8CEEA")
BORDER    = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

COL_LABELS = {
    "Numero da Oportunidade":    "Nr Oportunidade",
    "Tipo de Oportunidade":      "Tipo",
    "Criterio de Julgamento":    "Criterio",
    "Fim do periodo de cotacao": "Prazo",
    "Data (cotacao)":            "Data",
    "Hora (cotacao)":            "Hora",
    "Local de Entrega":          "Local",
    "Item":                      "Item",
    "Quantidade":                "Qtd",
    "Unidade de medida":         "Un",
    "Descricao de Item":         "Descricao Curta",
    "Descricao longa do item":   "Descricao Longa",
    "Fabricante/PN":             "Fabricante / PN",
    "Categoria":                 "Categoria",
    "Recorrente":                "Recorrente",
    "Responsavel":               "Responsavel",
}

COL_WIDTHS = {
    "Numero da Oportunidade":    16,
    "Tipo de Oportunidade":      22,
    "Criterio de Julgamento":    16,
    "Fim do periodo de cotacao": 22,
    "Data (cotacao)":            12,
    "Hora (cotacao)":            10,
    "Local de Entrega":          26,
    "Item":                       6,
    "Quantidade":                 8,
    "Unidade de medida":          8,
    "Descricao de Item":         32,
    "Descricao longa do item":   50,
    "Fabricante/PN":             22,
    "Categoria":                 16,
    "Recorrente":                12,
    "Responsavel":               14,
}

WRAP_COLS = {"Descricao de Item", "Descricao longa do item", "Local de Entrega"}


def _format_sheet(ws, df, title):
    ws.insert_rows(1, 2)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(df.columns))
    tc = ws.cell(1, 1, title)
    tc.font = Font(name="Arial", bold=True, color="FFFFFF", size=13)
    tc.fill = PatternFill("solid", fgColor="1237FF")
    tc.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28
    for ci, col in enumerate(df.columns, 1):
        c = ws.cell(2, ci, COL_LABELS.get(col, col))
        c.font = HDR_FONT; c.fill = HDR_FILL; c.alignment = CENTER; c.border = BORDER
    ws.row_dimensions[2].height = 22
    has_wrap = any(col in WRAP_COLS for col in df.columns)
    for ri, (_, row) in enumerate(df.iterrows(), 3):
        is_rec = str(row.get("Recorrente", "")).strip().lower() == "sim"
        fill = REC_FILL if is_rec else (ODD_FILL if ri % 2 else EVEN_FILL)
        for ci, col in enumerate(df.columns, 1):
            val = row[col]
            val = "" if pd.isna(val) else val
            c = ws.cell(ri, ci, val)
            c.font = BODY_FONT; c.fill = fill; c.border = BORDER
            c.alignment = LEFT if col in WRAP_COLS else CENTER
        ws.row_dimensions[ri].height = 40 if has_wrap else 18
    for ci, col in enumerate(df.columns, 1):
        ws.column_dimensions[get_column_letter(ci)].width = COL_WIDTHS.get(col, 18)
    ws.freeze_panes = "A3"


def _format_resumo(ws, resumo_df):
    ws.insert_rows(1, 2)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=3)
    tc = ws.cell(1, 1, "Resumo de Distribuicao")
    tc.font = Font(name="Arial", bold=True, color="FFFFFF", size=13)
    tc.fill = PatternFill("solid", fgColor="1237FF")
    tc.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28
    for ci, label in enumerate(["Responsavel", "Itens", "%"], 1):
        c = ws.cell(2, ci, label)
        c.font = HDR_FONT; c.fill = HDR_FILL; c.alignment = CENTER; c.border = BORDER
    ws.row_dimensions[2].height = 22
    for ri, (_, row) in enumerate(resumo_df.iterrows(), 3):
        fill = ODD_FILL if ri % 2 else EVEN_FILL
        for ci, val in enumerate([row["Responsavel"], row["Itens"], str(row["%"]) + "%"], 1):
            c = ws.cell(ri, ci, val)
            c.font = BODY_FONT; c.fill = fill; c.border = BORDER; c.alignment = CENTER
        ws.row_dimensions[ri].height = 18
    for col, w in zip("ABC", [20, 10, 10]):
        ws.column_dimensions[col].width = w


def to_excel(df):
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        if df.empty:
            pd.DataFrame({"Aviso": ["Nenhum item encontrado."]}).to_excel(writer, sheet_name="Consolidado", index=False)
            return buf.getvalue()
        for resp in RESPONSAVEIS:
            sub = df[df["Responsavel"] == resp]
            if not sub.empty:
                sub.to_excel(writer, sheet_name=resp[:31], index=False)
        df.to_excel(writer, sheet_name="Consolidado", index=False)
        resumo = pd.DataFrame([{
            "Responsavel": r,
            "Itens": int((df["Responsavel"] == r).sum()),
            "%": round((df["Responsavel"] == r).sum() / len(df) * 100, 1),
        } for r in RESPONSAVEIS])
        resumo.to_excel(writer, sheet_name="Resumo", index=False)

    wb = load_workbook(BytesIO(buf.getvalue()))
    for resp in RESPONSAVEIS:
        if resp in wb.sheetnames:
            sub = df[df["Responsavel"] == resp]
            _format_sheet(wb[resp], sub.reset_index(drop=True), "Next Supply - " + resp)
    if "Consolidado" in wb.sheetnames:
        _format_sheet(wb["Consolidado"], df.reset_index(drop=True), "Next Supply - Consolidado")
    if "Resumo" in wb.sheetnames:
        resumo = pd.DataFrame([{
            "Responsavel": r,
            "Itens": int((df["Responsavel"] == r).sum()),
            "%": round((df["Responsavel"] == r).sum() / len(df) * 100, 1),
        } for r in RESPONSAVEIS])
        _format_resumo(wb["Resumo"], resumo)

    out = BytesIO()
    wb.save(out)
    return out.getvalue()


# ── Interface ─────────────────────────────────────────────────────────────────

def render_topbar():
    st.markdown(
        '<div class="topbar"><div class="brand-wrap">'
        '<div class="brand-mark">N</div>'
        '<div><p class="brand-title">NEXT SUPPLY</p>'
        '<p class="brand-subtitle">Separador de OPS Petronect</p></div></div></div>',
        unsafe_allow_html=True
    )


def render_hero(df):
    total  = len(df) if df is not None and not df.empty else 0
    ops    = df["Numero da Oportunidade"].nunique() if df is not None and not df.empty else 0
    resp   = df["Responsavel"].nunique() if df is not None and not df.empty else 0
    status = "Processado" if total > 0 else "Aguardando upload"
    st.markdown(
        '<div class="hero"><div class="hero-title">Separador de OPS</div>'
        '<div class="hero-sub">Faca upload do ZIP exportado do Petronect. '
        'Para continuar o mes, carregue tambem o Excel mensal salvo anteriormente.</div>'
        '<div class="metrics">'
        '<div class="mcard"><div class="mcard-label">Oportunidades</div><div class="mcard-val accent">' + str(ops) + '</div></div>'
        '<div class="mcard"><div class="mcard-label">Itens extraidos</div><div class="mcard-val">' + str(total) + '</div></div>'
        '<div class="mcard"><div class="mcard-label">Responsaveis ativos</div><div class="mcard-val">' + str(resp) + '</div></div>'
        '<div class="mcard"><div class="mcard-label">Status</div><div class="mcard-val">' + status + '</div></div>'
        '</div></div>',
        unsafe_allow_html=True
    )


def sec(title):
    st.markdown('<div class="sec-title">' + title + '</div>', unsafe_allow_html=True)


def hr():
    st.markdown('<hr class="hr">', unsafe_allow_html=True)


# ── Main ──────────────────────────────────────────────────────────────────────

if "history" not in st.session_state:
    st.session_state["history"] = pd.DataFrame(columns=COLS)
if "last_upload" not in st.session_state:
    st.session_state["last_upload"] = None

render_topbar()

sec("Carregar historico mensal anterior (opcional)")
st.caption("Para continuar acumulando o mes e ativar deteccao de recorrencia, faca upload do Excel mensal salvo anteriormente.")

base_file = st.file_uploader("Excel mensal anterior (.xlsx)", type=["xlsx"], key="base_upload", label_visibility="collapsed")
if base_file:
    try:
        base_df = pd.read_excel(base_file, sheet_name="Consolidado", dtype=str)
        base_df.columns = [c.strip() for c in base_df.columns]
        inv_labels = {v: k for k, v in COL_LABELS.items()}
        base_df = base_df.rename(columns=inv_labels)
        for col in COLS:
            if col not in base_df.columns:
                base_df[col] = ""
        base_df = base_df[[c for c in COLS if c in base_df.columns]]
        for col in COLS:
            if col not in base_df.columns:
                base_df[col] = ""
        base_df = base_df[COLS]
        st.session_state["history"] = base_df.drop_duplicates(subset=["Numero da Oportunidade", "Item"], keep="last")
        st.success("Historico carregado: " + str(len(st.session_state["history"])) + " itens.")
    except Exception as e:
        st.error("Erro ao carregar base: " + str(e))

hr()

sec("Upload do arquivo ZIP")
st.caption("Selecione o arquivo compactado com os PDFs das oportunidades do dia.")
uploaded = st.file_uploader("Arquivo ZIP", type=["zip"], label_visibility="collapsed")
df_today = None

if uploaded and uploaded.name != st.session_state["last_upload"]:
    st.session_state["last_upload"] = uploaded.name
    with st.spinner("Processando PDFs..."):
        df_today, log = process_zip(uploaded.read(), st.session_state["history"])
    if df_today is not None and not df_today.empty:
        st.session_state["history"] = (
            pd.concat([st.session_state["history"], df_today], ignore_index=True)
            .drop_duplicates(subset=["Numero da Oportunidade", "Item"], keep="last")
        )
        n_rec = int((df_today["Recorrente"] == "Sim").sum())
        msg = "OK: " + str(len(df_today)) + " item(ns) de " + str(df_today["Numero da Oportunidade"].nunique()) + " oportunidade(s)."
        if n_rec > 0:
            msg += " | " + str(n_rec) + " item(ns) recorrente(s)."
        st.success(msg)
    else:
        st.warning("Nenhum item encontrado. Confira o log abaixo.")
    with st.expander("Log de processamento"):
        for line in log:
            st.write(line)

# Compatibilidade com historicos antigos
if not st.session_state["history"].empty:
    for col, default in [("Categoria", None), ("Recorrente", "Nao")]:
        if col not in st.session_state["history"].columns:
            if col == "Categoria":
                st.session_state["history"]["Categoria"] = st.session_state["history"].get(
                    "Descricao longa do item", pd.Series()
                ).apply(categorize)
            else:
                st.session_state["history"][col] = default

df_view = st.session_state["history"] if not st.session_state["history"].empty else None
render_hero(df_view)

if df_view is not None and not df_view.empty:
    hr()
    render_alertas(df_view)
    hr()

    sec("Filtros")
    c1, c2, c3, c4 = st.columns([2, 2, 2, 3])
    with c1:
        f_resp = st.multiselect("Responsavel", options=sorted(df_view["Responsavel"].dropna().unique()))
    with c2:
        cats = sorted(df_view["Categoria"].dropna().unique()) if "Categoria" in df_view.columns else []
        f_cat = st.multiselect("Categoria", options=cats)
    with c3:
        f_rec = st.selectbox("Recorrente", ["Todos", "Sim", "Nao"])
    with c4:
        f_text = st.text_input("Busca por descricao ou numero", "")

    df_filtered = df_view.copy()
    if f_resp:
        df_filtered = df_filtered[df_filtered["Responsavel"].isin(f_resp)]
    if f_cat:
        df_filtered = df_filtered[df_filtered["Categoria"].isin(f_cat)]
    if f_rec != "Todos":
        df_filtered = df_filtered[df_filtered["Recorrente"] == f_rec]
    if f_text:
        mask = (
            df_filtered["Descricao longa do item"].str.contains(f_text, case=False, na=False) |
            df_filtered["Numero da Oportunidade"].astype(str).str.contains(f_text, na=False)
        )
        df_filtered = df_filtered[mask]

    sec("Consolidado - " + str(len(df_filtered)) + " item(ns)")
    st.dataframe(df_filtered, use_container_width=True, height=380)
    hr()

    sec("Distribuicao")
    g1, g2, g3 = st.columns(3)
    with g1:
        st.caption("Por responsavel")
        dist_r = df_view["Responsavel"].value_counts().reset_index()
        dist_r.columns = ["Responsavel", "Itens"]
        st.bar_chart(dist_r.set_index("Responsavel"), color="#3F5DFF")
    with g2:
        st.caption("Por categoria")
        dist_c = df_view["Categoria"].value_counts().head(8).reset_index()
        dist_c.columns = ["Categoria", "Itens"]
        st.bar_chart(dist_c.set_index("Categoria"), color="#1237FF")
    with g3:
        st.caption("Por tipo de oportunidade")
        dist_t = df_view["Tipo de Oportunidade"].value_counts().head(8).reset_index()
        dist_t.columns = ["Tipo", "Itens"]
        st.bar_chart(dist_t.set_index("Tipo"), color="#7B9FFF")

    hr()
    sec("Exportar")
    e1, e2, e3 = st.columns(3)
    with e1:
        src = df_today if df_today is not None and not df_today.empty else df_view
        st.download_button(
            "Excel do dia",
            data=to_excel(src),
            file_name="nextsupply_ops_dia.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.caption("Somente o ultimo upload.")
    with e2:
        st.download_button(
            "Excel mensal (acumulado)",
            data=to_excel(st.session_state["history"]),
            file_name="nextsupply_ops_mensal.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.caption(str(len(st.session_state["history"])) + " itens acumulados. Salve para usar amanha.")
    with e3:
        if st.button("Limpar historico"):
            st.session_state["history"] = pd.DataFrame(columns=COLS)
            st.session_state["last_upload"] = None
            st.rerun()
        st.caption("Zera o acumulado para novo mes.")
