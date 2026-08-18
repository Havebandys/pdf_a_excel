from __future__ import annotations

import io
import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

import pandas as pd
import pdfplumber


DATE_FULL = re.compile(r"^\s*(\d{1,2}/\d{1,2}/\d{2,4})\s+")
MONEY = r"-?\d{1,3}(?:\.\d{3})*,\d{2}-?|-?\d+,\d{2}-?"


def ar_number(value: str | None) -> float | None:
    if not value or not value.strip():
        return None
    value = value.strip()
    trailing_minus = value.endswith("-")
    value = value.rstrip("-").replace("$", "").replace(" ", "")
    value = value.replace(".", "").replace(",", ".")
    try:
        number = float(value)
        return -abs(number) if trailing_minus else number
    except ValueError:
        return None


def pdf_to_layout_text(pdf_bytes: bytes) -> str:
    """Extracts positioned text in-process; no Poppler or external executable."""
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        return "\f".join(
            page.extract_text(layout=True, x_density=7.25, y_density=13) or ""
            for page in pdf.pages
        )


def detect_bank(text: str) -> str:
    top = text[:15000].upper()
    if "BANCO PATAGONIA" in top or "PATAGONIA E-BANK" in top:
        return "Patagonia"
    if "BBVA BANCO FRANC" in top or "PYMES Y NEGOCIOS" in top:
        return "BBVA"
    if "LISTADO DE MOVIMIENTOS HIST" in top and "COD TRX" in top:
        return "BTF"
    if "BANCO COMAFI" in top or "RESUMEN DE OPERACIONES" in top:
        return "Comafi"
    return "Desconocido"


def _date(value: str, default_year: int | None = None) -> pd.Timestamp | None:
    parts = value.split("/")
    if len(parts) == 2 and default_year:
        value = f"{value}/{default_year}"
    for fmt in ("%d/%m/%Y", "%d/%m/%y"):
        try:
            return pd.Timestamp(datetime.strptime(value, fmt))
        except ValueError:
            pass
    return None


def _year_near(text: str) -> int | None:
    period = re.search(r"movimientos del mes:\s*[A-ZÁÉÍÓÚ]+\s+(20\d{2})", text, re.I)
    if period:
        return int(period.group(1))
    hits = re.findall(r"(?:20\d{2}|\d{2})", text[:6000])
    years = []
    for hit in hits:
        year = int(hit)
        if year < 100:
            year += 2000
        if 2020 <= year <= 2035:
            years.append(year)
    return years[-1] if years else None


def parse_btf(text: str) -> tuple[list[dict], list[dict]]:
    rows, rejected = [], []
    pattern = re.compile(
        rf"^\s*(?P<date>\d{{2}}/\d{{2}}/\d{{4}})\s+(?P<origin>\d+)\s+"
        rf"(?P<body>.*?)\s+(?P<comp>\d+)\s+(?P<amount>{MONEY})\s+"
        rf"(?P<balance>{MONEY})\s+(?P<trx>\d+)\s*$"
    )
    for page_no, page in enumerate(text.split("\f"), 1):
        for line in page.splitlines():
            if not DATE_FULL.match(line):
                continue
            m = pattern.match(line)
            if not m:
                rejected.append({"Página": page_no, "Texto": line.strip(), "Motivo": "No se pudieron separar columnas"})
                continue
            amount = ar_number(m["amount"])
            rows.append({
                "Fecha": _date(m["date"]), "Operación": m["comp"],
                "Concepto": re.sub(r"\s+", " ", m["body"]).strip(),
                "Débito": abs(amount) if amount is not None and amount < 0 else None,
                "Crédito": amount if amount is not None and amount >= 0 else None,
                "Saldo": ar_number(m["balance"]), "Origen": m["origin"],
                "Código trx": m["trx"], "Página": page_no,
            })
    return rows, rejected


def _header_positions(page: str, bank: str):
    for line in page.splitlines():
        upper = line.upper()
        if "FECHA" in upper and "SALDO" in upper and ("DEBIT" in upper or "DÉBIT" in upper):
            pos = {
                "date": upper.find("FECHA"), "concept": upper.find("CONCEPTO"),
                "debit": max(upper.find("DEBIT"), upper.find("DÉBIT")),
                "credit": max(upper.find("CREDIT"), upper.find("CRÉDIT")),
                "balance": upper.rfind("SALDO"),
            }
            if bank == "Patagonia":
                pos["ref"] = upper.find("REFER")
                pos["value_date"] = upper.find("FECHA VALOR")
            elif bank == "BBVA":
                pos["origin"] = upper.find("ORIGEN")
            elif bank == "Comafi":
                pos["ref"] = upper.find("REFERENCIAS")
            return pos
    return None


def _slice(line: str, start: int, end: int | None) -> str:
    if start < 0:
        return ""
    return line[start:end].strip() if end is not None else line[start:].strip()


def parse_column_bank(text: str, bank: str) -> tuple[list[dict], list[dict]]:
    rows, rejected = [], []
    current_year = None
    current_account = ""
    for page_no, page in enumerate(text.split("\f"), 1):
        current_year = _year_near(page) or current_year
        account_matches = re.findall(r"(?:CUENTA CORRIENTE EN PESOS|CC \$|Cuenta Corriente Bancaria Nro\.)\s*([^\n]+)", page, re.I)
        if account_matches:
            current_account = re.sub(r"\s+", " ", account_matches[-1]).strip()[:80]
        pos = _header_positions(page, bank)
        if not pos:
            continue
        lines = page.splitlines()
        header_seen = False
        for line in lines:
            upper = line.upper()
            if "FECHA" in upper and "SALDO" in upper and ("DEBIT" in upper or "DÉBIT" in upper):
                header_seen = True
                continue
            if not header_seen:
                continue
            dm = re.match(r"^\s*(\d{1,2}/\d{1,2}(?:/\d{2,4})?)\s+", line)
            if not dm:
                continue
            date_text = dm.group(1)
            if date_text == "00/00" or "SIN MOVIMIENTOS" in upper or "SALDO ANTERIOR" in upper:
                continue
            date_val = _date(date_text, current_year)
            if date_val is None:
                rejected.append({"Página": page_no, "Texto": line.strip(), "Motivo": "Fecha no reconocida"})
                continue
            amounts = [(m, min(("debit", "credit", "balance"), key=lambda k: abs(m.end() - pos[k])))
                       for m in re.finditer(MONEY, line)]
            debit = next((ar_number(m.group()) for m, kind in amounts if kind == "debit"), None)
            credit = next((ar_number(m.group()) for m, kind in amounts if kind == "credit"), None)
            balance = next((ar_number(m.group()) for m, kind in amounts if kind == "balance"), None)
            c_start = dm.end()
            c_end = pos.get("ref", pos.get("debit", len(line)))
            if bank == "BBVA":
                c_start = max(c_start, pos.get("concept", c_start))
                c_end = amounts[0][0].start() if amounts else c_end
            concept = _slice(line, c_start, c_end)
            ref = _slice(line, pos.get("ref", -1), pos.get("value_date", pos["debit"]))
            origin = _slice(line, pos.get("origin", -1), pos["concept"])
            if debit is None and credit is None:
                rejected.append({"Página": page_no, "Texto": line.strip(), "Motivo": "Movimiento sin débito ni crédito"})
                continue
            rows.append({
                "Fecha": date_val, "Operación": ref, "Concepto": re.sub(r"\s+", " ", concept).strip(),
                "Débito": abs(debit) if debit is not None else None,
                "Crédito": abs(credit) if credit is not None else None,
                "Saldo": balance, "Origen": origin, "Código trx": "", "Página": page_no,
                "Cuenta": current_account,
            })
    return rows, rejected


def parse_comafi(text: str) -> tuple[list[dict], list[dict]]:
    """Comafi's text layer has stable numeric x-columns but wrapped credit rows."""
    rows, rejected = [], []
    current_account = ""
    for page_no, page in enumerate(text.split("\f"), 1):
        account = re.search(r"Cuenta Corriente Bancaria Nro\.\s*([\d-]+)", page, re.I)
        if account:
            current_account = account.group(1)
        if "DETALLE DE MOVIMIENTOS" not in page.upper():
            continue
        active = False
        last = None
        for line in page.splitlines():
            upper = line.upper()
            if "DETALLE DE MOVIMIENTOS" in upper:
                active = True
                continue
            if active and any(x in upper for x in ("IMPUESTOS DEBITADOS", "RESUMEN DE SALDO", "VISA DEBITO")):
                active = False
            if not active:
                continue
            dm = re.match(r"^\s*(\d{1,2}/\d{1,2}/\d{2,4})\s+(.*)$", line)
            money = list(re.finditer(MONEY, line))
            if dm:
                if "SALDO ANTERIOR" in upper or "SALDO AL:" in upper:
                    continue
                prefix_end = money[0].start() if money else len(line)
                prefix = line[dm.end(1):prefix_end].strip()
                refm = re.search(r"\s(\d{6,})\s*$", prefix)
                ref = refm.group(1) if refm else ""
                concept = prefix[:refm.start()].strip() if refm else prefix
                last = {
                    "Fecha": _date(dm.group(1)), "Operación": ref, "Concepto": re.sub(r"\s+", " ", concept),
                    "Débito": None, "Crédito": None, "Saldo": None, "Origen": "", "Código trx": "",
                    "Página": page_no, "Cuenta": current_account,
                }
                rows.append(last)
            elif last is not None and line.strip() and not money:
                extra = line.strip()
                if len(extra) < 100 and not any(x in upper for x in ("SALDO", "TOTAL", "FECHA")):
                    last["Concepto"] = (last["Concepto"] + " " + extra).strip()
            if last is not None and money:
                for m in money:
                    val = ar_number(m.group())
                    if m.start() >= 205:
                        last["Saldo"] = val
                    elif m.start() >= 175:
                        last["Crédito"] = abs(val) if val is not None else None
                    else:
                        last["Débito"] = abs(val) if val is not None else None
        rows = [r for r in rows if r["Débito"] is not None or r["Crédito"] is not None]
    return rows, rejected


def parse_pdf(pdf_bytes: bytes, forced_bank: str | None = None):
    text = pdf_to_layout_text(pdf_bytes)
    bank = forced_bank if forced_bank and forced_bank != "Automático" else detect_bank(text)
    if bank == "BTF":
        rows, rejected = parse_btf(text)
    elif bank in {"Patagonia", "BBVA"}:
        rows, rejected = parse_column_bank(text, bank)
    elif bank == "Comafi":
        rows, rejected = parse_comafi(text)
    else:
        rows, rejected = [], [{"Página": None, "Texto": "", "Motivo": "Banco no reconocido"}]
    df = pd.DataFrame(rows)
    if not df.empty:
        for col in ["Débito", "Crédito", "Saldo"]:
            df[col] = pd.to_numeric(df[col], errors="coerce")
        df.insert(0, "Banco", bank)
        df["Mes"] = df["Fecha"].dt.to_period("M").astype(str)
    return bank, df, pd.DataFrame(rejected)


def to_excel(bank: str, df: pd.DataFrame, rejected: pd.DataFrame) -> bytes:
    out = io.BytesIO()
    summary = pd.DataFrame({
        "Control": ["Banco", "Movimientos", "Total débitos", "Total créditos", "Filas a revisar"],
        "Resultado": [bank, len(df), df.get("Débito", pd.Series(dtype=float)).sum(),
                      df.get("Crédito", pd.Series(dtype=float)).sum(), len(rejected)],
    })
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Movimientos", index=False)
        summary.to_excel(writer, sheet_name="Control", index=False)
        rejected.to_excel(writer, sheet_name="Revisar", index=False)
        for ws in writer.book.worksheets:
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for col in ws.columns:
                letter = col[0].column_letter
                width = min(max((len(str(c.value or "")) for c in col), default=10) + 2, 60)
                ws.column_dimensions[letter].width = width
        if "Movimientos" in writer.book.sheetnames:
            ws = writer.book["Movimientos"]
            headers = {c.value: c.column for c in ws[1]}
            for name in ("Débito", "Crédito", "Saldo"):
                if name in headers:
                    for cell in ws.iter_cols(min_col=headers[name], max_col=headers[name], min_row=2):
                        for c in cell:
                            c.number_format = '#,##0.00;[Red]-#,##0.00'
            if "Fecha" in headers:
                for cell in ws.iter_cols(min_col=headers["Fecha"], max_col=headers["Fecha"], min_row=2):
                    for c in cell:
                        c.number_format = "dd/mm/yyyy"
    return out.getvalue()
