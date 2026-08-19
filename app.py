from __future__ import annotations

import base64
import io
import ipaddress
import re
import urllib.request
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st
from openpyxl.styles import Font, PatternFill
from supabase import Client, create_client

from parsers import parse_pdf


APP_VERSION = "Snoopy 2.0"
AUTHOR = "@PamperoSur"
TZ_AR = ZoneInfo("America/Argentina/Buenos_Aires")
MONTHS = ["", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
          "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
DISCLAIMER = (
    "Aplicación académica desarrollada en el marco de Emprendedurismo (IA) - Corrientes. "
    "Uso exclusivo de estudiantes y docentes autorizados. No constituye asesoramiento ni "
    "prestación de servicios contables, impositivos, financieros, jurídicos o profesionales. "
    "Los resultados deben verificarse contra la documentación bancaria original. Se prohíbe "
    "su uso o explotación comercial. El autor no asume responsabilidad por errores de "
    "extracción, interpretación o decisiones tomadas con la información procesada."
)

st.set_page_config(page_title="Snoopy 2.0 | PDF bancario → Excel", page_icon="🏦", layout="wide")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=IBM+Plex+Mono:wght@500;600&display=swap');
:root { --navy:#061525; --panel:#0a2139; --cyan:#48d7d0; --blue:#1597d5; --line:#245573; --muted:#9eb4c8; }
html, body, [class*="css"] { font-family:Inter,sans-serif; }
[data-testid="stAppViewContainer"] { background:
 linear-gradient(rgba(34,86,119,.055) 1px,transparent 1px),linear-gradient(90deg,rgba(34,86,119,.055) 1px,transparent 1px),
 radial-gradient(circle at 88% 3%,rgba(12,152,178,.18),transparent 25%),linear-gradient(145deg,#051321,#09213a 58%,#061725);
 background-size:38px 38px,38px 38px,auto,auto; color:#eef6ff; }
[data-testid="stSidebar"] { background:#061321; border-right:1px solid #1d4463; }
[data-testid="stHeader"] { background:transparent; }
.block-container { max-width:1500px; padding-top:1rem; }
.hero { display:grid; grid-template-columns:minmax(0,1fr) auto; gap:2rem; align-items:center; padding:1.6rem 1.75rem;
 border-radius:20px; border:1px solid #28617f; background:radial-gradient(circle at 83% 18%,rgba(24,184,190,.26),transparent 34%),linear-gradient(135deg,rgba(12,44,72,.98),rgba(7,27,47,.98));
 box-shadow:0 22px 60px rgba(0,0,0,.28),inset 0 1px 0 rgba(255,255,255,.05); margin-bottom:1rem; overflow:hidden; }
.hero h1 { margin:.28rem 0 .5rem; color:#fff; font-size:clamp(1.75rem,3vw,2.45rem); letter-spacing:-.035em; }
.hero p { margin:0; color:#b7cadb; max-width:900px; }
.eyebrow { color:#50d3c8; font-size:.76rem; font-weight:800; letter-spacing:.14em; text-transform:uppercase; }
.brand-lockup { min-width:225px; padding:.9rem 1.05rem; border:1px solid rgba(91,218,215,.35); border-radius:16px;
 background:rgba(4,20,35,.52); text-align:right; box-shadow:inset 0 0 28px rgba(31,181,188,.07); }
.brand-name { font:800 1.85rem 'IBM Plex Mono',monospace; letter-spacing:.055em;
 background:linear-gradient(90deg,#56e4d9,#56b9ff); -webkit-background-clip:text; background-clip:text; color:transparent; }
.brand-author { margin-top:.32rem; color:#fff; font:600 .76rem 'IBM Plex Mono',monospace; }
.brand-status { margin-top:.65rem; color:#72ddd6; font-size:.68rem; letter-spacing:.08em; text-transform:uppercase; }
.badge { display:inline-block; margin:.9rem .4rem 0 0; padding:.3rem .62rem; border-radius:999px; font-size:.72rem;
 background:#123451; border:1px solid #2b688e; color:#e3f3ff; }
.legal-title { color:#f2c65c; font:600 .72rem 'IBM Plex Mono',monospace; letter-spacing:.12em; text-transform:uppercase; margin-bottom:.45rem; }
.legal { margin-top:1.5rem; padding:1rem 1.2rem; border-radius:12px; border-left:4px solid #d6aa47;
 background:#13263c; color:#b7c7d8; font-size:.78rem; line-height:1.55; }
.author { color:#fff; font-weight:700; }
.status-grid { display:grid; grid-template-columns:repeat(4,minmax(0,1fr)); gap:.68rem; margin:.35rem 0 .85rem; }
.status-card { position:relative; padding:.82rem 1rem; border-radius:14px; border:1px solid rgba(52,110,144,.7);
 background:linear-gradient(145deg,rgba(13,43,68,.92),rgba(7,29,49,.92)); overflow:hidden; }
.status-card:before { content:""; position:absolute; left:0; top:0; bottom:0; width:3px; background:linear-gradient(#47ddd2,#1597d5); }
.status-label { color:#8faabd; font:600 .64rem 'IBM Plex Mono',monospace; letter-spacing:.1em; text-transform:uppercase; }
.status-value { margin-top:.28rem; color:#f7fbff; font-size:1.08rem; font-weight:750; }
.status-note { color:#6edfd7; font-size:.68rem; margin-top:.18rem; }
.kpi-grid { display:grid; grid-template-columns:repeat(4,minmax(0,1fr)); gap:.8rem; margin:1rem 0; }
.kpi { padding:1rem 1.05rem; border-radius:15px; border:1px solid #285a78; background:linear-gradient(145deg,#0d2b46,#081e33);
 box-shadow:0 12px 30px rgba(0,0,0,.16),inset 0 1px rgba(255,255,255,.04); }
.kpi-label { color:#94aec2; font:600 .65rem 'IBM Plex Mono',monospace; letter-spacing:.09em; text-transform:uppercase; }
.kpi-value { color:#fff; font-size:1.34rem; font-weight:800; margin-top:.35rem; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
.kpi-accent { height:2px; width:34px; margin-top:.65rem; background:linear-gradient(90deg,#4bdbd2,#168ec8); border-radius:3px; }
[data-testid="stFileUploader"] { background:#0b233a; border:1px dashed #37779d; border-radius:14px; padding:.5rem; }
.stTextInput input, [data-baseweb="select"] > div { border-radius:10px !important; }
.stTextInput input[aria-label="Usuario"] { text-transform:uppercase; letter-spacing:.045em; }
.login-shell { padding:.2rem .25rem .8rem; }
.login-kicker { color:#52d8cf; font:600 .68rem 'IBM Plex Mono',monospace; letter-spacing:.11em; text-transform:uppercase; }
.stColumn:has(.login-marker) [data-testid="stVerticalBlockBorderWrapper"] { min-height:430px; }
.excel-visual { position:relative; min-height:430px; height:100%; border-radius:18px; border:1px solid #285b79; overflow:hidden;
 background:radial-gradient(circle at 50% 38%,rgba(31,190,185,.17),transparent 48%),linear-gradient(145deg,#0c2943,#071c30); padding:1rem; }
.excel-visual:after { content:""; position:absolute; inset:0; background:linear-gradient(90deg,transparent 52%,rgba(7,28,48,.95) 96%); pointer-events:none; }
.doc-tag { color:#66ddd5; font:600 .66rem 'IBM Plex Mono',monospace; letter-spacing:.1em; text-transform:uppercase; }
.sheet { position:absolute; left:9%; top:18%; width:82%; transform:perspective(700px) rotateY(8deg) rotate(-1.5deg);
 border:1px solid rgba(111,213,209,.45); border-radius:10px; background:rgba(9,35,57,.88); box-shadow:0 24px 45px rgba(0,0,0,.34); overflow:hidden; }
.sheet-head,.sheet-row { display:grid; grid-template-columns:.7fr 1.65fr .8fr .8fr; }
.sheet-head span { padding:.55rem .42rem; background:#0e806f; color:#eafffd; font:600 .54rem 'IBM Plex Mono',monospace; }
.sheet-row span { padding:.48rem .42rem; border-right:1px solid rgba(56,105,131,.5); border-top:1px solid rgba(56,105,131,.42);
 color:#aec6d7; font:500 .52rem 'IBM Plex Mono',monospace; white-space:nowrap; overflow:hidden; }
.flow-arrow { position:absolute; right:5%; bottom:9%; color:#65e3d9; font:600 .66rem 'IBM Plex Mono',monospace; letter-spacing:.08em; z-index:2; }
.side-disclaimer { min-height:430px; height:100%; padding:.82rem .95rem; border-radius:18px; border:1px solid #665631; border-top:3px solid #e2b950;
 background:linear-gradient(165deg,rgba(40,47,54,.94),rgba(18,35,51,.96)); color:#afc0cf; font-size:.72rem; line-height:1.52; }
.side-disclaimer .legal-title { margin-bottom:.48rem; }
.side-disclaimer strong { color:#fff; }
.disclaimer-head { display:flex; align-items:flex-start; justify-content:space-between; gap:.65rem; margin-bottom:.45rem; }
.disclaimer-author { color:#e9c968; font:600 .62rem 'IBM Plex Mono',monospace; white-space:nowrap; text-align:right; }
.river-pride { margin-top:.62rem; padding-top:.58rem; border-top:1px solid rgba(222,188,88,.25); text-align:center; }
.river-pride img { display:block; width:82px; height:82px; object-fit:contain; margin:.05rem auto .25rem; filter:drop-shadow(0 8px 12px rgba(0,0,0,.28)); }
.river-title { color:#fff; font:800 1.05rem Inter,sans-serif; letter-spacing:.08em; }
.river-subtitle { margin-top:.08rem; color:#ed1b2f; font:700 .66rem 'IBM Plex Mono',monospace; letter-spacing:.16em; }
.world-banner { position:relative; margin-top:.5rem; height:66px; border-radius:16px; overflow:hidden; border:1px solid #2b6380;
 background:linear-gradient(180deg,rgba(101,190,225,.17) 0 33%,rgba(238,248,252,.10) 33% 66%,rgba(101,190,225,.17) 66%);
 box-shadow:inset 0 1px rgba(255,255,255,.05),0 12px 30px rgba(0,0,0,.14); }
.world-banner:before,.world-banner:after { content:""; position:absolute; top:0; bottom:0; width:90px; z-index:2; pointer-events:none; }
.world-banner:before { left:0; background:linear-gradient(90deg,#082039,transparent); }
.world-banner:after { right:0; background:linear-gradient(-90deg,#082039,transparent); }
.world-track { position:absolute; inset:0; display:flex; align-items:center; justify-content:center; animation:champions 8s ease-in-out infinite alternate; }
.champions { display:flex; align-items:center; gap:1.8rem; white-space:nowrap; }
.champion-title { color:#dcecf5; font:600 .68rem 'IBM Plex Mono',monospace; letter-spacing:.14em; }
.star-block { display:flex; align-items:center; gap:.45rem; color:#f0c75b; }
.star { font-size:1.75rem; line-height:1; filter:drop-shadow(0 0 7px rgba(240,199,91,.35)); }
.star-year { color:#f3d77f; font:600 .67rem 'IBM Plex Mono',monospace; }
@keyframes champions { from { transform:translateX(-9%); } to { transform:translateX(9%); } }
.stButton>button, .stDownloadButton>button { border-radius:9px; font-weight:700; min-height:2.65rem; }
.stButton>button[kind="primary"], .stDownloadButton>button[kind="primary"] {
 background:linear-gradient(90deg,#087ebc,#0da69c); color:#fff; border:0; }
hr { border-color:#20425d; }
@media(max-width:850px){.hero{grid-template-columns:1fr}.brand-lockup{text-align:left;min-width:0}.status-grid,.kpi-grid{grid-template-columns:1fr 1fr}.excel-visual,.side-disclaimer{min-height:340px}.world-track{animation:none}}
@media(max-width:560px){.status-grid,.kpi-grid{grid-template-columns:1fr}}
</style>
""", unsafe_allow_html=True)


def now_ar() -> datetime:
    return datetime.now(TZ_AR)


def period_label() -> str:
    value = now_ar()
    return f"{MONTHS[value.month]} {value.year}"


def hero(subtitle: str) -> None:
    st.markdown(f"""
    <div class="hero">
      <div><div class="eyebrow">Intelligence workspace · Uso educativo</div>
      <h1>PDF bancario → Excel normalizado</h1><p>{subtitle}</p></div>
      <div class="brand-lockup"><div class="brand-name">SNOOPY 2.0</div>
      <div class="brand-author">X: @PamperoSur · CAF</div>
      <div class="brand-status">{period_label()} · Corrientes</div></div>
    </div>""", unsafe_allow_html=True)


def money_ar(value: float) -> str:
    raw = f"{float(value):,.2f}"
    return "$ " + raw.replace(",", "X").replace(".", ",").replace("X", ".")


def datetime_ar(value) -> str:
    if not value or pd.isna(value):
        return "Sin ingresos"
    try:
        parsed = pd.to_datetime(value, utc=True).tz_convert(TZ_AR)
        return parsed.strftime("%d/%m/%Y %H:%M")
    except Exception:
        return str(value)


def clean_view(frame: pd.DataFrame) -> pd.DataFrame:
    view = frame.copy()
    if "Fecha" in view.columns:
        view["Fecha"] = pd.to_datetime(view["Fecha"], errors="coerce").dt.strftime("%d/%m/%Y")
    return view.where(pd.notna(view), "")


@st.cache_resource
def db() -> Client:
    return create_client(st.secrets["SUPABASE_URL"], st.secrets["SUPABASE_KEY"])


def headers() -> dict[str, str]:
    try:
        return {str(k).lower(): str(v) for k, v in st.context.headers.items()}
    except Exception:
        return {}


def client_info() -> dict[str, str]:
    h = headers()
    forwarded = h.get("x-forwarded-for", "")
    ip = forwarded.split(",")[0].strip() or h.get("x-real-ip", "N/D")
    ua = h.get("user-agent", "N/D")
    info = {"ip_address": ip, "city": h.get("x-vercel-ip-city", "N/D"),
            "region": h.get("x-vercel-ip-country-region", "N/D"),
            "country": h.get("cf-ipcountry", h.get("x-vercel-ip-country", "N/D")),
            "user_agent": ua}
    if info["city"] == "N/D" and ip not in {"", "N/D"}:
        geo = geolocate(ip)
        info.update({k: geo.get(k, info[k]) for k in ("city", "region", "country")})
    return info


@st.cache_data(ttl=3600, show_spinner=False)
def geolocate(ip: str) -> dict[str, str]:
    try:
        address = ipaddress.ip_address(ip)
        if address.is_private or address.is_loopback:
            return {}
        req = urllib.request.Request(f"https://ipwho.is/{ip}", headers={"User-Agent": "ExtractorBancario/2.0"})
        with urllib.request.urlopen(req, timeout=2) as response:
            import json
            data = json.loads(response.read().decode("utf-8"))
        if not data.get("success", True):
            return {}
        return {"city": data.get("city") or "N/D", "region": data.get("region") or "N/D",
                "country": data.get("country") or "N/D"}
    except Exception:
        return {}


def log_access(username: str, success: bool, event_type: str = "LOGIN") -> None:
    payload = {"username": username.lower().strip() or "(vacío)", "success": success,
               "event_type": event_type, **client_info()}
    try:
        db().table("access_logs").insert(payload).execute()
    except Exception:
        pass


def authenticate(username: str, password: str):
    result = db().rpc("verify_app_user", {"p_username": username, "p_password": password}).execute()
    rows = result.data or []
    if rows:
        log_access(username, True)
        return rows[0]
    log_access(username, False)
    return None


def academic_notice() -> None:
    st.markdown(f"""
    <div class="legal"><div class="legal-title">Descargo de responsabilidad</div>
    <span class="author">Autoría: {AUTHOR} · {period_label()}</span><br>
    <b>Emprendedurismo (IA) - Corrientes · {APP_VERSION} · USO EXCLUSIVAMENTE EDUCATIVO</b><br><br>
    {DISCLAIMER}</div>""", unsafe_allow_html=True)


@st.cache_data(show_spinner=False)
def river_logo_data_url() -> str:
    logo_path = Path(__file__).with_name("river_logo.png")
    encoded = base64.b64encode(logo_path.read_bytes()).decode("ascii")
    return f"data:image/png;base64,{encoded}"


def excel_preview() -> None:
    st.markdown("""
    <div class="excel-visual"><div class="doc-tag">PDF RAW DATA → EXCEL</div>
      <div class="sheet">
        <div class="sheet-head"><span>FECHA</span><span>CONCEPTO</span><span>CRÉDITO</span><span>DÉBITO</span></div>
        <div class="sheet-row"><span>03/08</span><span>TRANSFERENCIA CUIT 20...</span><span>125.000</span><span>—</span></div>
        <div class="sheet-row"><span>05/08</span><span>DEPÓSITO EFECTIVO</span><span>82.500</span><span>—</span></div>
        <div class="sheet-row"><span>11/08</span><span>PAGO PROVEEDOR</span><span>—</span><span>43.200</span></div>
        <div class="sheet-row"><span>18/08</span><span>TRANSFERENCIA RECIBIDA</span><span>310.800</span><span>—</span></div>
        <div class="sheet-row"><span>22/08</span><span>IMPUESTO DÉB./CRÉD.</span><span>—</span><span>1.865</span></div>
      </div><div class="flow-arrow">NORMALIZADO ✓</div>
    </div>""", unsafe_allow_html=True)


def side_disclaimer() -> None:
    st.markdown(f"""
    <div class="side-disclaimer"><div class="disclaimer-head">
      <div class="legal-title">Descargo de responsabilidad</div><div class="disclaimer-author">AUTORÍA · {AUTHOR}</div></div>
      <strong>Uso exclusivamente educativo.</strong><br><br>{DISCLAIMER}
      <div class="river-pride"><img src="{river_logo_data_url()}" alt="Escudo de River Plate">
      <div class="river-title">EL MÁS GRANDE</div><div class="river-subtitle">ORGULLO MILLONARIO</div></div>
    </div>""", unsafe_allow_html=True)


def champions_banner() -> None:
    st.markdown("""
    <div class="world-banner"><div class="world-track"><div class="champions">
      <div class="champion-title">ARGENTINA · TRES VECES CAMPEÓN DEL MUNDO</div>
      <div class="star-block"><span class="star">★</span><span class="star-year">1978</span></div>
      <div class="star-block"><span class="star">★</span><span class="star-year">1986</span></div>
      <div class="star-block"><span class="star">★</span><span class="star-year">2022</span></div>
      <div class="champion-title">ORGULLO · HISTORIA · IDENTIDAD</div>
    </div></div></div>""", unsafe_allow_html=True)


def login_screen() -> None:
    hero("Extractor y herramienta de uso contable académico. Emprendedurismo (IA) · Corrientes.")
    st.markdown(f"""<div class="status-grid">
      <div class="status-card"><div class="status-label">Cobertura</div><div class="status-value">7 entidades bancarias</div><div class="status-note">Lectores normalizados</div></div>
      <div class="status-card"><div class="status-label">Privacidad</div><div class="status-value">Procesamiento temporal</div><div class="status-note">Los PDF no se almacenan</div></div>
      <div class="status-card"><div class="status-label">Robustez</div><div class="status-value">PDF de gran volumen</div><div class="status-note">Procesamiento página por página</div></div>
      <div class="status-card"><div class="status-label">Fiabilidad</div><div class="status-value">Salida normalizada</div><div class="status-note">Excel, CSV y control de filas</div></div>
    </div>""", unsafe_allow_html=True)
    visual, center, legal = st.columns(3, gap="medium")
    with visual:
        excel_preview()
    with center:
        with st.container(border=True):
            st.markdown('<span class="login-marker"></span>', unsafe_allow_html=True)
            st.markdown('<div class="login-kicker">Acceso seguro · Control de usuarios</div>', unsafe_allow_html=True)
            st.subheader("Ingreso de usuarios")
            username = st.text_input("Usuario").strip().lower()
            password = st.text_input("Clave", type="password")
            accepted = st.checkbox(
                "He leído y acepto el uso exclusivamente educativo, el descargo de responsabilidad y el registro de acceso."
            )
            if st.button("Ingresar", type="primary", width="stretch", disabled=not accepted):
                try:
                    user = authenticate(username, password)
                    if user:
                        st.session_state.user = user
                        st.rerun()
                    else:
                        st.error("Usuario, clave o estado incorrecto.")
                except Exception as exc:
                    st.error(f"No fue posible conectar con la base de usuarios: {exc}")
    with legal:
        side_disclaimer()
    champions_banner()


def classify_origin(text: str) -> str:
    value = str(text).upper()
    rules = [("DEBIN", "DEBIN"), ("DEPOS", "Depósito"), ("EFECTIVO", "Efectivo"),
             ("CHEQUE", "Cheque"), ("POSNET", "Cobranza/POSNET"), ("TARJ", "Tarjeta"),
             ("TRANSF", "Transferencia"), ("COELS", "Transferencia"),
             ("TR INTER", "Transferencia"), ("CRED TR", "Transferencia")]
    return next((label for token, label in rules if token in value), "Otro/N.D.")


def fiscalize(frame: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    df = frame.copy()
    df["Año"] = df["Fecha"].dt.year
    df["Mes número"] = df["Fecha"].dt.month
    df["Mes"] = df["Mes número"].map(lambda x: MONTHS[int(x)] if pd.notna(x) else "")
    df["Año-Mes"] = df["Fecha"].dt.strftime("%Y-%m")
    df["CUIT/DNI detectado"] = df["Concepto"].astype(str).str.extract(r"(?<!\d)(\d{11})(?!\d)", expand=False).fillna("N/D")
    df["Procedencia"] = df["Concepto"].map(classify_origin)
    df["Banco receptor"] = df["Banco"]
    df["Banco origen"] = "N/D"
    df["Nombre/Procedencia detectada"] = df.apply(extract_name, axis=1)
    preferred = ["Banco", "Fecha", "Concepto", "Crédito", "Débito", "CUIT/DNI detectado",
                 "Nombre/Procedencia detectada", "Procedencia", "Banco receptor", "Banco origen",
                 "Operación", "Página", "Año", "Mes", "Año-Mes", "Origen", "Código trx", "Saldo"]
    df = df[[c for c in preferred if c in df.columns] + [c for c in df.columns if c not in preferred]]
    credits = df[pd.to_numeric(df["Crédito"], errors="coerce").fillna(0) > 0].copy()
    monthly = credits.groupby("Año-Mes", as_index=False).agg(
        Acreditaciones=("Crédito", "size"), Total_acreditado=("Crédito", "sum")
    )
    return df, credits, monthly


def extract_name(row) -> str:
    text = re.sub(r"\s+", " ", str(row.get("Concepto", ""))).strip()
    identifier = str(row.get("CUIT/DNI detectado", "N/D"))
    if identifier != "N/D" and identifier in text:
        tail = text.split(identifier, 1)[1].strip(" -:/")
        return tail[:90] if tail else "N/D"
    return "N/D"


def export_workbook(bank: str, full: pd.DataFrame, credits: pd.DataFrame,
                    grouped: pd.DataFrame, monthly: pd.DataFrame, rejected: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    control = pd.DataFrame({"Control": ["Aplicación", "Finalidad", "Autoría", "Entorno", "Emisión",
                                                "Banco", "Movimientos", "Acreditaciones", "Total acreditado", "Filas a revisar"],
                            "Resultado": [f"Extractor Bancario IA - {APP_VERSION}", "Uso exclusivamente educativo",
                                          AUTHOR, "Emprendedurismo (IA) - Corrientes", period_label(), bank,
                                          len(full), len(credits), credits["Crédito"].sum(), len(rejected)]})
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        full.to_excel(writer, sheet_name="Movimientos", index=False)
        credits.to_excel(writer, sheet_name="Créditos", index=False)
        grouped.to_excel(writer, sheet_name="Agrupación créditos", index=False)
        monthly.to_excel(writer, sheet_name="Resumen mensual", index=False)
        control.to_excel(writer, sheet_name="Control", index=False)
        rejected.to_excel(writer, sheet_name="Revisar", index=False)
        for ws in writer.book.worksheets:
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for cell in ws[1]:
                cell.fill = PatternFill("solid", fgColor="123A59")
                cell.font = Font(color="FFFFFF", bold=True)
            for col in ws.columns:
                ws.column_dimensions[col[0].column_letter].width = min(
                    max(len(str(c.value or "")) for c in col) + 2, 55)
            headers = {cell.value: cell.column for cell in ws[1]}
            if "Fecha" in headers:
                for column in ws.iter_cols(min_col=headers["Fecha"], max_col=headers["Fecha"], min_row=2):
                    for cell in column:
                        cell.number_format = "dd/mm/yyyy"
            for name in ("Crédito", "Débito", "Saldo", "Total_acreditado"):
                if name in headers:
                    for column in ws.iter_cols(min_col=headers[name], max_col=headers[name], min_row=2):
                        for cell in column:
                            cell.number_format = '#,##0.00;[Red]-#,##0.00'
    return output.getvalue()


def extractor_page() -> None:
    hero("Conversión prioritaria y análisis fiscalizador de acreditaciones, CUIT, fecha, monto y procedencia.")
    if "conversions_session" not in st.session_state:
        st.session_state.conversions_session = 0
    st.markdown(f"""<div class="status-grid">
      <div class="status-card"><div class="status-label">Sesión actual</div><div class="status-value">{st.session_state.conversions_session} PDF procesados</div><div class="status-note">Sin persistencia documental</div></div>
      <div class="status-card"><div class="status-label">Seguridad</div><div class="status-value">Usuario autenticado</div><div class="status-note">Acceso y actividad registrados</div></div>
      <div class="status-card"><div class="status-label">Período operativo</div><div class="status-value">{period_label()}</div><div class="status-note">Actualización mensual automática</div></div>
    </div>""", unsafe_allow_html=True)
    bank_choice = st.selectbox("Banco / lector", ["Automático", "BTF", "Patagonia", "BBVA", "Comafi", "Macro", "Galicia", "HSBC"])
    uploaded = st.file_uploader("Seleccionar un extracto PDF", type=["pdf"], accept_multiple_files=False)
    upload_token = (uploaded.name, uploaded.size) if uploaded else None
    previous_token = st.session_state.get("active_upload")
    if previous_token and upload_token != previous_token:
        st.session_state.pop("result", None)
        st.session_state.pop("downloads", None)
        st.session_state.pop("active_upload", None)
        if uploaded is None:
            st.toast("PDF retirado: datos y descargas eliminados de la sesión.", icon="🧹")
    if uploaded and st.button("Convertir extracto", type="primary", width="stretch"):
        with st.spinner("Extrayendo y normalizando movimientos página por página…"):
            try:
                pdf_bytes = uploaded.getvalue()
                if len(pdf_bytes) > 60 * 1024 * 1024:
                    raise ValueError("El archivo supera el límite operativo de 60 MB.")
                progress_bar = st.progress(0, text="Preparando el PDF…")
                def update_progress(done: int, total: int) -> None:
                    progress_bar.progress(done / total, text=f"Procesando página {done} de {total}")
                bank, frame, rejected = parse_pdf(pdf_bytes, bank_choice, update_progress)
                progress_bar.progress(1.0, text="Extracción terminada")
                if frame.empty:
                    raise ValueError(f"No se detectaron movimientos para el lector {bank}. Revisá la pestaña Control o elegí el banco manualmente.")
                full, credits, monthly = fiscalize(frame)
                grouped = credits.groupby(
                    ["CUIT/DNI detectado", "Nombre/Procedencia detectada", "Procedencia", "Banco receptor"],
                    dropna=False, as_index=False
                ).agg(Acreditaciones=("Crédito", "size"), Total_acreditado=("Crédito", "sum")).sort_values("Total_acreditado", ascending=False)
                st.session_state.result = (uploaded.name, bank, full, credits, grouped, monthly, rejected)
                st.session_state.active_upload = upload_token
                st.session_state.conversions_session += 1
                st.session_state.pop("downloads", None)
                log_access(st.session_state.user["username"], True, "PDF_PROCESSED")
            except MemoryError:
                st.error("El servidor agotó memoria. El proceso se detuvo de forma controlada; probá dividir el PDF por períodos.")
            except Exception as exc:
                st.error(f"No se pudo convertir el PDF: {exc}")
    if "result" not in st.session_state:
        st.info("Subí un PDF. El documento y sus movimientos no se guardan en la base de usuarios.")
        academic_notice()
        return
    _, bank, full, credits, grouped, monthly, rejected = st.session_state.result
    total_credits = credits["Crédito"].sum()
    movement_count = f"{len(full):,}".replace(",", ".")
    credit_count = f"{len(credits):,}".replace(",", ".")
    st.markdown(f"""<div class="kpi-grid">
      <div class="kpi"><div class="kpi-label">Entidad detectada</div><div class="kpi-value">{bank}</div><div class="kpi-accent"></div></div>
      <div class="kpi"><div class="kpi-label">Movimientos</div><div class="kpi-value">{movement_count}</div><div class="kpi-accent"></div></div>
      <div class="kpi"><div class="kpi-label">Acreditaciones</div><div class="kpi-value">{credit_count}</div><div class="kpi-accent"></div></div>
      <div class="kpi"><div class="kpi-label">Total acreditado</div><div class="kpi-value" title="{money_ar(total_credits)}">{money_ar(total_credits)}</div><div class="kpi-accent"></div></div>
    </div>""", unsafe_allow_html=True)
    tabs = st.tabs(["Movimientos", "Fiscalización de créditos", "Agrupaciones", "Control", "Descargas"])
    with tabs[0]:
        preview = full.head(2000)
        if len(full) > len(preview):
            st.info(f"Vista optimizada: se muestran 2.000 de {len(full):,} movimientos. El Excel contiene el total.")
        st.dataframe(clean_view(preview), hide_index=True, width="stretch")
    with tabs[1]:
        col1, col2, col3 = st.columns(3)
        min_amount = col1.number_input("Crédito mínimo", min_value=0.0, value=0.0, step=10000.0)
        search = col2.text_input("Buscar CUIT, nombre o concepto")
        origins = col3.multiselect("Procedencia", sorted(credits["Procedencia"].dropna().unique()))
        view = credits[credits["Crédito"] >= min_amount]
        if search:
            mask = view.astype(str).apply(lambda col: col.str.contains(search, case=False, na=False)).any(axis=1)
            view = view[mask]
        if origins:
            view = view[view["Procedencia"].isin(origins)]
        ordered = view.sort_values("Crédito", ascending=False)
        if len(ordered) > 2000:
            st.info(f"Se muestran las primeras 2.000 de {len(ordered):,} acreditaciones filtradas.")
        st.dataframe(clean_view(ordered.head(2000)), hide_index=True, width="stretch")
    with tabs[2]:
        st.markdown("#### Por CUIT, persona, procedencia y banco receptor")
        st.dataframe(clean_view(grouped), hide_index=True, width="stretch")
        st.markdown("#### Por mes")
        st.dataframe(clean_view(monthly), hide_index=True, width="stretch")
    with tabs[3]:
        if rejected.empty:
            st.success("No se detectaron líneas dudosas.")
        else:
            st.warning(f"Hay {len(rejected)} líneas que requieren control.")
            st.dataframe(rejected, hide_index=True, width="stretch")
    with tabs[4]:
        st.caption("El archivo se prepara únicamente cuando lo solicitás, para reducir el uso de memoria del servidor.")
        if st.button("Preparar archivos de descarga", type="primary", width="stretch"):
            with st.spinner("Generando Excel y CSV…"):
                st.session_state.downloads = (
                    export_workbook(bank, full, credits, grouped, monthly, rejected),
                    credits.to_csv(index=False, sep=";", decimal=",", date_format="%d/%m/%Y").encode("utf-8-sig"),
                )
        if "downloads" in st.session_state:
            book, csv = st.session_state.downloads
            d1, d2 = st.columns(2)
            d1.download_button("Descargar Excel fiscalizador", book, f"{bank.lower()}_fiscalizacion.xlsx",
                               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary", width="stretch")
            d2.download_button("Descargar créditos CSV", csv, f"{bank.lower()}_creditos.csv", "text/csv", width="stretch")
    academic_notice()


def admin_users_page() -> None:
    st.title("Administración de usuarios")
    st.caption("Disponible exclusivamente para Administradores.")
    with st.expander("Crear usuario", expanded=True):
        with st.form("create_user", clear_on_submit=True):
            c1, c2 = st.columns(2)
            username = c1.text_input("Usuario").strip().lower()
            full_name = c2.text_input("Nombre completo").strip()
            password = c1.text_input("Clave inicial", type="password")
            role = c2.selectbox("Rol", ["Analista", "Administrador"])
            if st.form_submit_button("Crear usuario", type="primary", width="stretch"):
                try:
                    db().rpc("create_app_user", {"p_username": username, "p_full_name": full_name,
                                                  "p_password": password, "p_role": role,
                                                  "p_created_by": st.session_state.user["username"]}).execute()
                    st.success(f"Usuario {username} creado.")
                    st.rerun()
                except Exception as exc:
                    st.error(f"No se pudo crear: {exc}")
    users = db().table("app_users").select(
        "username,full_name,role,active,created_at,created_by,last_login,access_count"
    ).order("username").execute().data or []
    user_df = pd.DataFrame(users).rename(columns={
        "username": "Usuario", "full_name": "Nombre completo", "role": "Rol", "active": "Activo",
        "created_at": "Creado", "created_by": "Creado por", "last_login": "Último ingreso",
        "access_count": "Cantidad de accesos"})
    if not user_df.empty:
        user_df["Creado"] = user_df["Creado"].map(datetime_ar)
        user_df["Último ingreso"] = user_df["Último ingreso"].map(datetime_ar)
    st.dataframe(user_df, hide_index=True, width="stretch")
    editable = [u["username"] for u in users if u["username"] != "adm"]
    if editable:
        st.markdown("#### Modificar usuario")
        target = st.selectbox("Usuario", editable)
        c1, c2, c3 = st.columns(3)
        new_role = c1.selectbox("Nuevo rol", ["Analista", "Administrador"], key="new_role")
        if c1.button("Cambiar rol", width="stretch"):
            db().table("app_users").update({"role": new_role}).eq("username", target).execute(); st.rerun()
        new_password = c2.text_input("Nueva clave", type="password")
        if c2.button("Restablecer clave", width="stretch", disabled=not new_password):
            db().rpc("reset_app_password", {"p_username": target, "p_new_password": new_password}).execute(); st.success("Clave actualizada.")
        selected = next(u for u in users if u["username"] == target)
        if c3.button("Desactivar" if selected["active"] else "Activar", width="stretch"):
            db().table("app_users").update({"active": not selected["active"]}).eq("username", target).execute(); st.rerun()
        if c3.button("Eliminar usuario", width="stretch"):
            db().table("app_users").delete().eq("username", target).execute(); st.rerun()


def access_page() -> None:
    st.title("Registro de accesos")
    logs = db().table("access_logs").select(
        "username,event_type,success,ip_address,city,region,country,user_agent,created_at"
    ).order("created_at", desc=True).limit(1000).execute().data or []
    frame = pd.DataFrame(logs)
    if not frame.empty:
        frame = frame.rename(columns={"username": "Usuario", "event_type": "Evento", "success": "Correcto",
                                             "ip_address": "IP", "city": "Ciudad", "region": "Región",
                                             "country": "País", "user_agent": "Navegador/dispositivo",
                                             "created_at": "Fecha y hora"})
        frame["Fecha y hora"] = frame["Fecha y hora"].map(datetime_ar)
    st.dataframe(frame, hide_index=True, width="stretch")
    if not frame.empty:
        st.download_button("Descargar registro CSV", frame.to_csv(index=False).encode("utf-8-sig"),
                           "registro_accesos.csv", "text/csv")


if "user" not in st.session_state:
    login_screen()
    st.stop()

user = st.session_state.user
with st.sidebar:
    st.markdown("### Snoopy 2.0")
    st.caption(f"Extractor Bancario IA · {AUTHOR}")
    st.caption(f"{user['full_name']} · {user['role']}")
    pages = ["Extractor"]
    if user["role"] == "Administrador":
        pages += ["Usuarios", "Accesos"]
    page = st.radio("Navegación", pages)
    st.divider()
    st.caption(f"Último acceso: {datetime_ar(user.get('last_login'))} · Argentina")
    if st.button("Cerrar sesión", width="stretch"):
        log_access(user["username"], True, "LOGOUT")
        st.session_state.clear()
        st.rerun()

if page == "Extractor":
    extractor_page()
elif page == "Usuarios":
    admin_users_page()
else:
    access_page()
