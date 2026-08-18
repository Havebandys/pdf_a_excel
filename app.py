from __future__ import annotations

import io
import ipaddress
import re
import urllib.request
from datetime import datetime
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st
from openpyxl.styles import Font, PatternFill
from supabase import Client, create_client

from parsers import parse_pdf


APP_VERSION = "Snoopy 2.0"
AUTHOR = "@CAF"
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
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
html, body, [class*="css"] { font-family:Inter,sans-serif; }
[data-testid="stAppViewContainer"] { background:linear-gradient(145deg,#061525,#0a2139 55%,#071827); color:#eef6ff; }
[data-testid="stSidebar"] { background:#061321; border-right:1px solid #1d4463; }
[data-testid="stHeader"] { background:transparent; }
.block-container { max-width:1500px; padding-top:1.25rem; }
.hero { padding:1.55rem 1.75rem; border-radius:18px; border:1px solid #235577;
 background:radial-gradient(circle at 87% 18%,rgba(16,163,184,.30),transparent 32%),linear-gradient(135deg,#0b2842,#0a1d32);
 box-shadow:0 18px 55px rgba(0,0,0,.27); margin-bottom:1rem; }
.hero h1 { margin:.25rem 0 .45rem; color:#fff; font-size:2.05rem; }
.hero p { margin:0; color:#b7cadb; max-width:950px; }
.eyebrow { color:#50d3c8; font-size:.76rem; font-weight:800; letter-spacing:.14em; text-transform:uppercase; }
.badge { display:inline-block; margin:.9rem .4rem 0 0; padding:.3rem .62rem; border-radius:999px; font-size:.72rem;
 background:#123451; border:1px solid #2b688e; color:#e3f3ff; }
.legal { margin-top:1.5rem; padding:1rem 1.2rem; border-radius:12px; border-left:4px solid #d6aa47;
 background:#13263c; color:#b7c7d8; font-size:.78rem; line-height:1.55; }
.author { color:#fff; font-weight:700; }
[data-testid="stMetric"] { background:#0c263f; border:1px solid #235274; border-radius:14px; padding:.8rem 1rem; }
[data-testid="stFileUploader"] { background:#0b233a; border:1px dashed #37779d; border-radius:14px; padding:.5rem; }
.stButton>button, .stDownloadButton>button { border-radius:9px; font-weight:700; min-height:2.65rem; }
.stButton>button[kind="primary"], .stDownloadButton>button[kind="primary"] {
 background:linear-gradient(90deg,#087ebc,#0da69c); color:#fff; border:0; }
hr { border-color:#20425d; }
</style>
""", unsafe_allow_html=True)


def now_ar() -> datetime:
    return datetime.now(TZ_AR)


def period_label() -> str:
    value = now_ar()
    return f"{MONTHS[value.month]} {value.year}"


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
    <div class="legal"><span class="author">Autoría: {AUTHOR} · {period_label()}</span><br>
    <b>Emprendedurismo (IA) - Corrientes · {APP_VERSION} · USO EXCLUSIVAMENTE EDUCATIVO</b><br><br>
    {DISCLAIMER}</div>""", unsafe_allow_html=True)


def login_screen() -> None:
    st.markdown(f"""
    <div class="hero"><div class="eyebrow">{APP_VERSION} · Uso exclusivamente educativo</div>
    <h1>PDF bancario → Excel normalizado</h1>
    <p>Extractor y herramienta de fiscalización académica. Emprendedurismo (IA) - Corrientes.</p></div>
    """, unsafe_allow_html=True)
    _, center, _ = st.columns([1, 1.15, 1])
    with center:
        with st.container(border=True):
            st.subheader("Ingreso de usuarios")
            username = st.text_input("Usuario").strip().lower()
            password = st.text_input("Clave", type="password")
            accepted = st.checkbox(
                "He leído y acepto el uso exclusivamente educativo, el descargo de responsabilidad y el registro de acceso."
            )
            st.caption("Se registra usuario, fecha, hora, IP, información técnica y ubicación aproximada con fines de seguridad y administración académica.")
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
    academic_notice()


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
    st.markdown(f"""
    <div class="hero"><div class="eyebrow">{APP_VERSION} · Uso exclusivamente educativo</div>
    <h1>PDF bancario → Excel normalizado</h1>
    <p>Conversión prioritaria y análisis fiscalizador de acreditaciones, CUIT, fecha, monto y procedencia.</p>
    <span class="badge">BTF</span><span class="badge">Patagonia</span><span class="badge">BBVA</span><span class="badge">Comafi</span>
    <span class="badge">Macro</span><span class="badge">Galicia</span><span class="badge">HSBC</span></div>
    """, unsafe_allow_html=True)
    bank_choice = st.selectbox("Banco / lector", ["Automático", "BTF", "Patagonia", "BBVA", "Comafi", "Macro", "Galicia", "HSBC"])
    uploaded = st.file_uploader("Seleccionar un extracto PDF", type=["pdf"], accept_multiple_files=False)
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
                st.session_state.pop("downloads", None)
            except MemoryError:
                st.error("El servidor agotó memoria. El proceso se detuvo de forma controlada; probá dividir el PDF por períodos.")
            except Exception as exc:
                st.error(f"No se pudo convertir el PDF: {exc}")
    if "result" not in st.session_state:
        st.info("Subí un PDF. El documento y sus movimientos no se guardan en la base de usuarios.")
        academic_notice()
        return
    _, bank, full, credits, grouped, monthly, rejected = st.session_state.result
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Banco", bank)
    c2.metric("Movimientos", f"{len(full):,}".replace(",", "."))
    c3.metric("Acreditaciones", f"{len(credits):,}".replace(",", "."))
    total_credits = credits["Crédito"].sum()
    c4.metric("Total acreditado", money_ar(total_credits), help=f"Importe exacto: {money_ar(total_credits)}")
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
    st.caption("Extractor Bancario IA · @CAF")
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
