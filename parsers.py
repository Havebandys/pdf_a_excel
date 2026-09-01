from __future__ import annotations

import gc
import io
import re
from datetime import datetime
from typing import Callable

import pandas as pd
import pdfplumber

AR_MONEY = r"-?\d{1,3}(?:\.\d{3})*,\d{2}-?|-?\d+,\d{2}-?"
US_MONEY = r"-?(?:\d{1,3}(?:,\d{3})+|\d+)\.\d{2}-?"
SANTANDER_MONEY = re.compile(r"-?\$\s*(\d{1,3}(?:\.\d{3})*,\d{2}|\d+,\d{2})")
FULL_DATE = re.compile(r"^\s*(\d{1,2}/\d{1,2}/\d{2,4})\s+")
MONTHS_EN = {"JAN": 1, "FEB": 2, "MAR": 3, "APR": 4, "MAY": 5, "JUN": 6,
             "JUL": 7, "AUG": 8, "SEP": 9, "OCT": 10, "NOV": 11, "DEC": 12}


def ar_number(value: str | None) -> float | None:
    if not value or not value.strip():
        return None
    raw = value.strip().replace("$", "").replace(" ", "")
    trailing_minus = raw.endswith("-")
    raw = raw.rstrip("-").replace(".", "").replace(",", ".")
    try:
        number = float(raw)
        return -abs(number) if trailing_minus else number
    except ValueError:
        return None


def us_number(value: str | None) -> float | None:
    if not value or not value.strip():
        return None
    raw = value.strip().replace("$", "").replace(" ", "")
    trailing_minus = raw.endswith("-")
    raw = raw.rstrip("-").replace(",", "")
    try:
        number = float(raw)
        return -abs(number) if trailing_minus else number
    except ValueError:
        return None


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


def _hsbc_date(value: str, year: int | None) -> pd.Timestamp | None:
    match = re.match(r"(\d{1,2})-([A-Z]{3})", value.upper())
    if not match or not year or match.group(2) not in MONTHS_EN:
        return None
    return pd.Timestamp(year=year, month=MONTHS_EN[match.group(2)], day=int(match.group(1)))


def _year_near(text: str) -> int | None:
    patterns = [
        r"EXTRACTO\s+DEL\s+\d{1,2}/\d{1,2}/(20\d{2})",
        r"PER[IÍ]ODO\s+DE\s+MOVIMIENTOS[\s\S]{0,180}?(20\d{2})",
        r"F:\d{1,2}/\d{1,2}/(\d{2})",
        r"MOVIMIENTOS\s+DEL\s+MES:\s*[A-ZÁÉÍÓÚ]+\s+(20\d{2})",
    ]
    for pattern in patterns:
        match = re.search(pattern, text, re.I)
        if match:
            year = int(match.group(1))
            return year + 2000 if year < 100 else year
    years = [int(y) for y in re.findall(r"\b(20(?:2\d|3[0-5]))\b", text[:7000])]
    return years[-1] if years else None


def detect_bank(text: str) -> str:
    top = re.sub(r"\s+", "", text[:30000].upper())
    if "RESUMENDECUENTA" in top and ("BANCOSANTANDERARGENTINA" in top or "SANTANDER" in top):
        return "Santander"
    if "FECHATRX" in top and "IMPORTEMO" in top and "SALDO_PROMEDIO" in top:
        return "Macro"
    if "BANCOGALICIA" in top or "RESUMENDECUENTACORRIENTEENPESOS" in top:
        return "Galicia"
    if "EXTRACTODEL" in top and "ESTIMADOSSE" in top and "REFERENCIA" in top:
        return "HSBC"
    if "BANCOPATAGONIA" in top or "PATAGONIAE-BANK" in top:
        return "Patagonia"
    # Some BBVA PDFs expose the vertical legal footer backwards ("AVBB")
    # instead of extracting "Banco BBVA Argentina" normally.  The customer
    # service and movement-section markers remain readable and are specific
    # enough to identify the statement safely.
    if ("BBVABANCOFRANC" in top or "PYMESYNEGOCIOS" in top
            or "BANCOBBVAARGENTINA" in top or "LINEABBVA" in top
            or "WWW.BBVA.COM.AR" in top
            or ("MOVIMIENTOSENCUENTAS" in top and "CTA.CTE.BANCARIA" in top)):
        return "BBVA"
    if (("LISTADODEMOVIMIENTOSHIST" in top and "CODTRX" in top)
            or ("LIQUIDACIONDEPRESENTACIONDECUPONES" in top
                and "BANCODETIERRADELFUEGO" in top)):
        return "BTF"
    if "BANCOCOMAFI" in top or "RESUMENDEOPERACIONES" in top:
        return "Comafi"
    return "Desconocido"


def _account(page: str, bank: str, previous: str = "") -> str:
    patterns = {
        "Macro": r"^\s*\d{1,2}/\d{1,2}/\d{4}\s+(\d{10,})",
        "Galicia": r"(?:N[ÚU]MERO DE CUENTA|CUENTA:)\s*(?:N[°º]\s*)?([\d-]+)",
        "BBVA": r"CC\s*\$\s*([\d-]+/\d)",
        "HSBC": r"CUENTA CORRIENTE EN \$ NRO\.\s*([\d-]+)",
        "Patagonia": r"CUENTA CORRIENTE EN PESOS\s*([^\n]+)",
        "Comafi": r"Cuenta Corriente Bancaria Nro\.\s*([\d-]+)",
        "Santander": r"Cuenta Corriente (?:en pesos )?N[º°]\s*([\d-]+(?:/\d+)?)",
        "BTF": r"(?:CUENTA:[^\n]{0,120}?NRO:|NRO:)\s*([\d-]{6,})",
    }
    match = re.search(patterns.get(bank, r"$^"), page, re.I | re.M)
    if not match:
        generic = r"(?:CUENTA|CTA)\s*(?:CORRIENTE|CAJA\s+DE\s+AHORRO)?\s*(?:EN\s+(?:PESOS|\$))?\s*(?:N(?:RO)?[.°º]*|N[ÚU]MERO)?\s*[:#-]?\s*([\d-]{6,}(?:/\d+)?)"
        match = re.search(generic, page, re.I | re.M)
    return re.sub(r"\s+", " ", match.group(1)).strip()[:80] if match else previous


def _bbva_section_account(line: str) -> str | None:
    """Return the account only for a real BBVA movement-section heading.

    BBVA continuation pages can start with movements from the account opened on
    the preceding page and introduce another account farther down.  Therefore a
    page-wide search cannot safely decide the account for every row on the page.
    """
    match = re.match(
        r"^\s*CC\s*(?:U\$S|\$)\s*([\d-]+/\d+)\s*"
        r"\(\s*Cta\.?\s*Cte\.?\s*Bancaria\s*\)",
        line,
        re.I,
    )
    return match.group(1).strip() if match else None


def _header_positions(page: str, bank: str) -> dict[str, int] | None:
    for line in page.splitlines():
        upper = line.upper()
        if "FECHA" in upper and "SALDO" in upper and ("DEBIT" in upper or "DÉBIT" in upper):
            pos = {"date": upper.find("FECHA"), "concept": upper.find("CONCEPTO"),
                   "debit": max(upper.find("DEBIT"), upper.find("DÉBIT")),
                   "credit": max(upper.find("CREDIT"), upper.find("CRÉDIT")),
                   "balance": upper.rfind("SALDO")}
            if bank == "Galicia":
                pos["concept"] = upper.find("DESCRIP")
                pos["origin"] = upper.find("ORIGEN")
            elif bank == "Patagonia":
                pos["ref"] = upper.find("REFER")
                pos["value_date"] = upper.find("FECHA VALOR")
            elif bank == "BBVA":
                pos["origin"] = upper.find("ORIGEN")
            return pos
    return None


def _parse_column_page(page: str, bank: str, page_no: int, state: dict) -> tuple[list[dict], list[dict]]:
    rows, rejected = [], []
    state["year"] = _year_near(page) or state.get("year")
    # BBVA must keep the account inherited from the preceding page until a real
    # movement-section heading is reached while reading from top to bottom.
    # Other banks retain their established page-wide account detection.
    if bank != "BBVA":
        state["account"] = _account(page, bank, state.get("account", ""))
    pos = _header_positions(page, bank)
    if not pos:
        return rows, rejected
    active = False
    for line in page.splitlines():
        upper = line.upper()
        if bank == "BBVA":
            section_account = _bbva_section_account(line)
            if section_account:
                previous_account = state.get("account", "")
                state["account"] = section_account
                if section_account != previous_account:
                    state.setdefault("account_transitions", []).append({
                        "Página": page_no,
                        "Cuenta anterior": previous_account or "N/D",
                        "Cuenta nueva": section_account,
                    })
                # A new account section must expose its own column header before
                # any following date-like line can be interpreted as a movement.
                active = False
                continue
        if "FECHA" in upper and "SALDO" in upper and ("DEBIT" in upper or "DÉBIT" in upper):
            active = True
            continue
        if not active:
            continue
        match = re.match(r"^\s*(\d{1,2}/\d{1,2}(?:/\d{2,4})?)\s+", line)
        if not match or "SALDO ANTERIOR" in upper:
            continue
        date_value = _date(match.group(1), state.get("year"))
        if date_value is None:
            rejected.append({"Página": page_no, "Texto": line.strip(), "Motivo": "Fecha no reconocida"})
            continue
        money = list(re.finditer(AR_MONEY, line))
        assigned = [(item, min(("debit", "credit", "balance"), key=lambda key: abs(item.start() - pos[key]))) for item in money]
        debit = next((ar_number(item.group()) for item, kind in assigned if kind == "debit"), None)
        credit = next((ar_number(item.group()) for item, kind in assigned if kind == "credit"), None)
        balance = next((ar_number(item.group()) for item, kind in assigned if kind == "balance"), None)
        if debit is None and credit is None:
            rejected.append({"Página": page_no, "Texto": line.strip(), "Motivo": "Movimiento sin débito ni crédito"})
            continue
        first_amount = min((item.start() for item in money), default=len(line))
        concept_start = max(match.end(), pos.get("concept", match.end()))
        concept = line[concept_start:first_amount].strip()
        if bank == "Galicia" and pos.get("origin", -1) > concept_start:
            concept = line[concept_start:pos["origin"]].strip()
        origin = ""
        if pos.get("origin", -1) >= 0:
            origin = line[pos["origin"]:min(first_amount, pos["credit"])].strip()
        ref = ""
        if pos.get("ref", -1) >= 0:
            ref = line[pos["ref"]:pos.get("value_date", pos["debit"])].strip()
        rows.append({"Fecha": date_value, "Operación": ref, "Concepto": re.sub(r"\s+", " ", concept),
                     "Débito": abs(debit) if debit is not None else None,
                     "Crédito": abs(credit) if credit is not None else None,
                     "Saldo": balance, "Origen": origin, "Código trx": "", "Página": page_no,
                     "Cuenta": state.get("account", "")})
    return rows, rejected


def _parse_macro_page(page: str, page_no: int, state: dict) -> tuple[list[dict], list[dict]]:
    rows, rejected = [], []
    pattern = re.compile(r"^\s*(?P<date>\d{1,2}/\d{1,2}/\d{4})\s+(?P<account>\d{10,})\s*"
                         r"(?P<body>.*?)\s+(?P<branch>\d{1,4})\s+(?P<amount>-?\d+(?:\.\d+)?)\s*"
                         r"(?P<tail>.*?)\s*(?P<balance>-?\d+(?:\.\d+)?)\s+(?P<average>-?\d+(?:\.\d+)?)\s*$")
    for line in page.splitlines():
        if not FULL_DATE.match(line):
            continue
        match = pattern.match(line)
        if not match:
            rejected.append({"Página": page_no, "Texto": line.strip(), "Motivo": "Columnas Macro no reconocidas"})
            continue
        amount = float(match.group("amount"))
        body = re.sub(r"\s+", " ", match.group("body")).strip()
        tail = re.sub(r"\s+", " ", match.group("tail")).strip(" -")
        rows.append({"Fecha": _date(match.group("date")), "Operación": body.split(" - ", 1)[0],
                     "Concepto": f"{body} {tail}".strip(), "Débito": abs(amount) if amount < 0 else None,
                     "Crédito": amount if amount >= 0 else None, "Saldo": float(match.group("balance")),
                     "Origen": match.group("branch"), "Código trx": "", "Página": page_no,
                     "Cuenta": match.group("account")})
    return rows, rejected


def _parse_macro_page_object(page, page_no: int, state: dict) -> tuple[list[dict], list[dict]]:
    """Reads Macro by physical x-columns because amount and reference touch in its text layer."""
    rows, rejected = [], []
    date_words = [word for word in page.extract_words() if re.fullmatch(r"\d{1,2}/\d{1,2}/\d{4}", word["text"])]
    for word in date_words:
        top, bottom = max(0, word["top"] - 1), min(page.height, word["bottom"] + 1)
        def column(x0: float, x1: float) -> str:
            return (page.crop((x0, top, x1, bottom)).extract_text(x_tolerance=1, y_tolerance=1) or "").strip()
        meta = column(55, 350)
        branch = column(350, 380)
        amount_text = column(380, 418).replace(" ", "").rstrip("-")
        detail = column(418, 640)
        numeric_detail = re.match(r"\d", detail)
        if "." in amount_text:
            whole, decimals = amount_text.split(".", 1)
            amount_text = whole + "." + decimals[:2]
        elif numeric_detail and amount_text and amount_text[-1] == detail[0]:
            amount_text = amount_text[:-1]
        balance_text = column(640, 675).replace(" ", "")
        average_text = column(675, 760).replace(" ", "")
        meta_match = re.match(r"(\d{1,2}/\d{1,2}/\d{4})\s+(\d{15})(.*)", meta)
        amount_match = re.match(r"-?\d+(?:\.\d+)?", amount_text)
        try:
            amount = float(amount_match.group()) if amount_match else float("nan")
            balance = float(balance_text)
            if pd.isna(amount):
                raise ValueError
        except ValueError:
            rejected.append({"Página": page_no, "Texto": " | ".join((meta, branch, amount_text, detail, balance_text, average_text)),
                             "Motivo": "Columnas físicas Macro no reconocidas"})
            continue
        if not meta_match:
            rejected.append({"Página": page_no, "Texto": meta, "Motivo": "Fecha/cuenta Macro no reconocidas"})
            continue
        body = re.sub(r"\s+", " ", meta_match.group(3)).strip()
        rows.append({"Fecha": _date(meta_match.group(1)), "Operación": body.split(" - ", 1)[0],
                     "Concepto": re.sub(r"\s+", " ", f"{body} {detail}").strip(" -"),
                     "Débito": abs(amount) if amount < 0 else None,
                     "Crédito": amount if amount >= 0 else None, "Saldo": balance,
                     "Origen": (re.match(r"\d+", branch).group() if re.match(r"\d+", branch) else branch),
                     "Código trx": "", "Página": page_no,
                     "Cuenta": meta_match.group(2)})
    return rows, rejected


def _parse_hsbc_page(page: str, page_no: int, state: dict) -> tuple[list[dict], list[dict]]:
    rows, rejected = [], []
    state["year"] = _year_near(page) or state.get("year")
    state["account"] = _account(page, "HSBC", state.get("account", ""))
    pos = _header_positions(page, "HSBC")
    if not pos:
        return rows, rejected
    for line in page.splitlines():
        match = re.match(r"^\s*(\d{1,2}-[A-Z]{3})\s+(.*)$", line, re.I)
        if match:
            state["current_date"] = _hsbc_date(match.group(1), state.get("year"))
            content_start = match.end(1)
        else:
            content_start = 0
        date_value = state.get("current_date")
        money = list(re.finditer(US_MONEY, line))
        candidate = line[content_start:].strip()
        if date_value is None or not money or not candidate.startswith("-"):
            continue
        assigned = [(item, min(("debit", "credit", "balance"), key=lambda key: abs(item.start() - pos[key]))) for item in money]
        debit = next((us_number(item.group()) for item, kind in assigned if kind == "debit"), None)
        credit = next((us_number(item.group()) for item, kind in assigned if kind == "credit"), None)
        balance = next((us_number(item.group()) for item, kind in assigned if kind == "balance"), None)
        first_amount = min(item.start() for item in money)
        prefix = line[content_start:first_amount].strip()
        ref_match = re.search(r"\s(\d{4,})\s*$", prefix)
        operation = ref_match.group(1) if ref_match else ""
        concept = prefix[:ref_match.start()].strip(" -") if ref_match else prefix.strip(" -")
        rows.append({"Fecha": date_value, "Operación": operation, "Concepto": re.sub(r"\s+", " ", concept),
                     "Débito": abs(debit) if debit is not None else None, "Crédito": abs(credit) if credit is not None else None,
                     "Saldo": balance, "Origen": "", "Código trx": "", "Página": page_no,
                     "Cuenta": state.get("account", "")})
    return rows, rejected


def _parse_galicia_page(page: str, page_no: int, state: dict) -> tuple[list[dict], list[dict]]:
    """Galicia's embedded font joins words; signed amounts still identify debit/credit reliably."""
    rows, rejected = [], []
    state["account"] = _account(page, "Galicia", state.get("account", ""))
    for line in page.splitlines():
        match = re.match(r"^(\d{1,2}/\d{1,2}/\d{2,4})\s+(.*)$", line.strip())
        if not match:
            continue
        date_value = _date(match.group(1))
        money = list(re.finditer(AR_MONEY, line))
        if date_value is None or len(money) < 2:
            rejected.append({"Página": page_no, "Texto": line.strip(), "Motivo": "Fila Galicia no reconocida"})
            continue
        amount = ar_number(money[-2].group())
        balance = ar_number(money[-1].group())
        concept = line[match.end(1):money[-2].start()].strip()
        rows.append({"Fecha": date_value, "Operación": "", "Concepto": re.sub(r"\s+", " ", concept),
                     "Débito": abs(amount) if amount is not None and amount < 0 else None,
                     "Crédito": amount if amount is not None and amount >= 0 else None,
                     "Saldo": balance, "Origen": "", "Código trx": "", "Página": page_no,
                     "Cuenta": state.get("account", "")})
    return rows, rejected


def _parse_btf_page(page: str, page_no: int, state: dict) -> tuple[list[dict], list[dict]]:
    rows, rejected = [], []
    state["account"] = _account(page, "BTF", state.get("account", ""))
    pattern = re.compile(rf"^\s*(?P<date>\d{{2}}/\d{{2}}/\d{{4}})\s+(?P<origin>\d+)\s+"
                         rf"(?P<body>.*?)\s+(?P<comp>\d+)\s+(?P<amount>{AR_MONEY})\s+"
                         rf"(?P<balance>{AR_MONEY})\s+(?P<trx>\d+)\s*$")
    for line in page.splitlines():
        if not FULL_DATE.match(line):
            continue
        # BTF repeats the report emission date in the page header.  It is not a
        # movement and must not inflate the "Revisar" sheet.
        if re.fullmatch(r"\s*\d{1,2}/\d{1,2}/\d{4}\s*", line):
            continue
        match = pattern.match(line)
        if not match:
            rejected.append({"Página": page_no, "Texto": line.strip(), "Motivo": "Columnas BTF no reconocidas"})
            continue
        amount = ar_number(match.group("amount"))
        rows.append({"Fecha": _date(match.group("date")), "Operación": match.group("comp"),
                     "Concepto": re.sub(r"\s+", " ", match.group("body")).strip(),
                     "Débito": abs(amount) if amount is not None and amount < 0 else None,
                     "Crédito": amount if amount is not None and amount >= 0 else None,
                     "Saldo": ar_number(match.group("balance")), "Origen": match.group("origin"),
                     "Código trx": match.group("trx"), "Página": page_no,
                     "Cuenta": state.get("account", "")})
    return rows, rejected


def _parse_btf_liquidation_page(page: str, page_no: int, state: dict) -> tuple[list[dict], list[dict]]:
    """Parse one BTF merchant-card settlement as one bank accreditation.

    These PDFs are not current-account statements: every page contains a
    settlement plus a duplicate receipt.  Reading the transaction table would
    duplicate amounts, so the authoritative values are the payment date,
    settlement number, destination CBU and final net payment shown in the
    settlement summary.
    """
    rows, rejected = [], []
    if "LIQUIDACION DE PRESENTACION DE CUPONES" not in page.upper():
        return rows, rejected

    date_match = re.search(r"Fecha de pago:\s*(\d{1,2}/\d{1,2}/\d{4})", page, re.I)
    liquidation_match = re.search(r"N\s*Liquidaci\S*:\s*([\d.]+)", page, re.I)
    cbu_match = re.search(r"ACRED\s+EN\s+CBU\s+(\d{18,24})", page, re.I)
    net_matches = re.findall(
        rf"IMPORTE\s+NETO\s+DE\s+PAGOS\s+\$\s*({AR_MONEY})", page, re.I
    )

    missing = []
    if not date_match:
        missing.append("fecha de pago")
    if not liquidation_match:
        missing.append("número de liquidación")
    if not cbu_match:
        missing.append("CBU")
    if not net_matches:
        missing.append("importe neto")
    if missing:
        rejected.append({
            "Página": page_no,
            "Texto": "Liquidación BTF",
            "Motivo": f"Faltan campos: {', '.join(missing)}",
        })
        return rows, rejected

    amount = ar_number(net_matches[0])
    if amount is None:
        rejected.append({
            "Página": page_no,
            "Texto": net_matches[0],
            "Motivo": "Importe neto BTF no reconocido",
        })
        return rows, rejected

    account = cbu_match.group(1)
    state["account"] = account
    rows.append({
        "Fecha": _date(date_match.group(1)),
        "Operación": liquidation_match.group(1),
        "Concepto": "LIQUIDACION DE TARJETAS A COMERCIOS",
        "Débito": abs(amount) if amount < 0 else None,
        "Crédito": amount if amount >= 0 else None,
        "Saldo": None,
        "Origen": "BTF TARJETAS",
        "Código trx": "",
        "Página": page_no,
        "Cuenta": account,
    })
    return rows, rejected


def _parse_comafi_page(page: str, page_no: int, state: dict) -> tuple[list[dict], list[dict]]:
    """Parse Comafi statements using the columns printed in each page header.

    ``extract_text(layout=True)`` produces lines much shorter than the original
    PDF coordinates, so fixed character offsets are not reliable.  The column
    centres below are derived from the local ``Débitos / Créditos / Saldo``
    header and therefore work across the different Comafi layouts in the same
    historical PDF.
    """
    rows, rejected = [], []
    active, last = False, None
    column_centres = state.get("comafi_column_centres")

    for line in page.splitlines():
        upper = line.upper()

        # A statement can begin a second account halfway down the same page.
        # Only this real section heading is allowed to change the active account;
        # operation references must never be interpreted as account numbers.
        account_match = re.search(r"^\s*N[ÚU]MERO\s+([\d-]+)\s+CBU\s*:", line, re.I)
        if account_match:
            state["account"] = account_match.group(1)
            active = False
            last = None
            continue

        if "DETALLE DE MOVIMIENTOS" in upper:
            active = True
            continue

        # Continuation pages do not always repeat DETALLE DE MOVIMIENTOS, but
        # they do repeat the actual movement-column header.
        if all(token in upper for token in ("FECHA", "CONCEPT", "DÉBITOS", "CRÉDITOS", "SALDO")) \
                or all(token in upper for token in ("FECHA", "CONCEPT", "DEBITOS", "CREDITOS", "SALDO")):
            debit_start = upper.find("DÉBITOS") if "DÉBITOS" in upper else upper.find("DEBITOS")
            credit_start = upper.find("CRÉDITOS") if "CRÉDITOS" in upper else upper.find("CREDITOS")
            balance_start = upper.rfind("SALDO")
            column_centres = {
                "Débito": debit_start + 3.5,
                "Crédito": credit_start + 4.0,
                "Saldo": balance_start + 2.5,
            }
            state["comafi_column_centres"] = column_centres
            active = bool(state.get("account"))
            last = None
            continue

        if active and any(token in upper for token in ("IMPUESTOS DEBITADOS", "RESUMEN DE SALDO", "VISA DEBITO")):
            active = False
            last = None
        if not active:
            continue
        date_match = re.match(r"^\s*(\d{1,2}/\d{1,2}/\d{2,4})\s+(.*)$", line)
        money = list(re.finditer(AR_MONEY, line))
        if date_match and "SALDO ANTERIOR" not in upper and "SALDO AL:" not in upper:
            prefix_end = money[0].start() if money else len(line)
            prefix = line[date_match.end(1):prefix_end].strip()
            ref_match = re.search(r"\s(\d{6,})\s*$", prefix)
            last = {"Fecha": _date(date_match.group(1)), "Operación": ref_match.group(1) if ref_match else "",
                    "Concepto": prefix[:ref_match.start()].strip() if ref_match else prefix,
                    "Débito": None, "Crédito": None, "Saldo": None, "Origen": "", "Código trx": "",
                    "Página": page_no, "Cuenta": state.get("account", "")}
            rows.append(last)
        if last is not None and money and column_centres:
            for item in money:
                value = ar_number(item.group())
                item_centre = (item.start() + item.end()) / 2
                column = min(column_centres, key=lambda name: abs(item_centre - column_centres[name]))
                if column == "Saldo":
                    last["Saldo"] = value
                elif column == "Crédito":
                    last["Crédito"] = abs(value) if value is not None else None
                else:
                    last["Débito"] = abs(value) if value is not None else None
    return [row for row in rows if row["Débito"] is not None or row["Crédito"] is not None], rejected


def _parse_comafi_page_object(page_obj, page_no: int, state: dict) -> tuple[list[dict], list[dict]]:
    """Parse Comafi from positioned PDF words instead of character offsets."""
    rows, rejected = [], []
    words = page_obj.extract_words(use_text_flow=False, keep_blank_chars=False) or []
    grouped: list[list[dict]] = []
    for word in sorted(words, key=lambda item: (round(item["top"], 1), item["x0"])):
        if not grouped or abs(grouped[-1][0]["top"] - word["top"]) > 1.5:
            grouped.append([word])
        else:
            grouped[-1].append(word)

    active = False
    # Some Comafi transactions continue on the following page: the dated line
    # is at the bottom of one page and the beneficiary/amount is at the top of
    # the next.  Keep the same row object in state so that amount is not lost.
    last = state.get("comafi_last")
    column_centres = state.get("comafi_x_centres")
    for line_words in grouped:
        line_words.sort(key=lambda item: item["x0"])
        text = " ".join(item["text"] for item in line_words)
        upper = text.upper()

        account_match = re.search(r"\bN[ÚU]MERO\s+([\d-]+)\s+CBU\s*:", text, re.I)
        if account_match:
            state["account"] = account_match.group(1)
            active = False
            last = None
            state["comafi_last"] = None
            continue

        if "DETALLE DE MOVIMIENTOS" in upper:
            active = True
            continue

        header_words = {item["text"].upper(): item for item in line_words}
        debit_header = header_words.get("DÉBITOS") or header_words.get("DEBITOS")
        credit_header = header_words.get("CRÉDITOS") or header_words.get("CREDITOS")
        balance_header = header_words.get("SALDO")
        if "FECHA" in header_words and debit_header and credit_header and balance_header:
            column_centres = {
                "Débito": (debit_header["x0"] + debit_header["x1"]) / 2,
                "Crédito": (credit_header["x0"] + credit_header["x1"]) / 2,
                "Saldo": (balance_header["x0"] + balance_header["x1"]) / 2,
            }
            state["comafi_x_centres"] = column_centres
            active = bool(state.get("account"))
            continue

        if active and any(token in upper for token in (
                "IMPUESTOS DEBITADOS", "RESUMEN DE SALDO", "VISA DEBITO")):
            active = False
            last = None
            state["comafi_last"] = None
        if not active or not column_centres:
            continue

        date_match = re.match(r"^(\d{1,2}/\d{1,2}/\d{2,4})\b", text)
        money_words = [item for item in line_words if re.fullmatch(AR_MONEY, item["text"])]
        if "TRANSPORTE" in upper:
            # Transporte is a page carry-forward marker, not the closing balance
            # of the pending transaction printed immediately before it.
            continue
        if date_match and "SALDO ANTERIOR" not in upper and "SALDO AL:" not in upper:
            first_money_x = min((item["x0"] for item in money_words), default=float("inf"))
            prefix_words = [item["text"] for item in line_words
                            if item["x0"] < first_money_x and item["text"] != date_match.group(1)]
            prefix = " ".join(prefix_words).strip()
            ref_match = re.search(r"\s(\d{6,})\s*$", prefix)
            last = {
                "Fecha": _date(date_match.group(1)),
                "Operación": ref_match.group(1) if ref_match else "",
                "Concepto": prefix[:ref_match.start()].strip() if ref_match else prefix,
                "Débito": None,
                "Crédito": None,
                "Saldo": None,
                "Origen": "",
                "Código trx": "",
                "Página": page_no,
                "Cuenta": state.get("account", ""),
            }
            rows.append(last)
            state["comafi_last"] = last

        if last is not None:
            if money_words and last not in rows and all(
                    last.get(name) is None for name in ("Débito", "Crédito")):
                rows.append(last)
            for item in money_words:
                value = ar_number(item["text"])
                item_centre = (item["x0"] + item["x1"]) / 2
                column = min(column_centres, key=lambda name: abs(item_centre - column_centres[name]))
                if column == "Saldo":
                    last["Saldo"] = value
                elif column == "Crédito":
                    last["Crédito"] = abs(value) if value is not None else None
                else:
                    last["Débito"] = abs(value) if value is not None else None

    state["comafi_last"] = last

    return [row for row in rows if row["Débito"] is not None or row["Crédito"] is not None], rejected


def _parse_santander_page(page: str, page_no: int, state: dict) -> tuple[list[dict], list[dict]]:
    """Parse Santander summaries while excluding tax recap and legal pages."""
    rows, rejected = [], []
    state["account"] = _account(page, "Santander", state.get("account", ""))
    upper_page = page.upper()
    if "MOVIMIENTOS" not in upper_page and "FECHA COMPROBANTE MOVIMIENTO" not in upper_page:
        return rows, rejected

    active = False
    pending_date = None
    last = None
    for line in page.splitlines():
        clean = line.strip()
        upper = clean.upper()
        if "FECHA" in upper and "COMPROBANTE" in upper and "MOVIMIENTO" in upper and "SALDO" in upper:
            active = True
            continue
        if active and any(token in upper for token in ("DETALLE IMPOSITIVO", "SALDO TOTAL", "BANCO SANTANDER ARGENTINA S.A.")):
            active = False
        if not active or not clean:
            continue

        date_match = re.match(r"^(\d{1,2}/\d{1,2}/\d{2,4})(?:\s+(.*))?$", clean)
        if date_match:
            pending_date = _date(date_match.group(1))
            remainder = (date_match.group(2) or "").strip()
        else:
            remainder = clean

        amounts = list(SANTANDER_MONEY.finditer(line))
        if not amounts:
            if remainder and last is not None and not re.match(r"^(?:BANCO SANTANDER|\* SALVO|\d+\s*-\s*\d+)", upper):
                last["Concepto"] = re.sub(r"\s+", " ", f'{last["Concepto"]} {remainder}').strip()
            continue
        if pending_date is None or "SALDO INICIAL" in upper or upper.startswith("RESPONSABLE:"):
            continue
        if len(amounts) < 2:
            # Supporting lines may contain a taxable base but are not movements.
            if last is not None and upper.startswith("RESPONSABLE:"):
                last["Concepto"] = re.sub(r"\s+", " ", f'{last["Concepto"]} {remainder}').strip()
            continue

        transaction_amount, balance_amount = amounts[-2], amounts[-1]
        amount = ar_number(transaction_amount.group())
        balance = ar_number(balance_amount.group())
        if transaction_amount.group().lstrip().startswith("-"):
            amount = -abs(amount) if amount is not None else None
        if balance_amount.group().lstrip().startswith("-"):
            balance = -abs(balance) if balance is not None else None

        prefix_start = (line.find(date_match.group(1)) + len(date_match.group(1))) if date_match else 0
        prefix = line[prefix_start:transaction_amount.start()].strip()
        operation_match = re.match(r"(\d+)\s+(.*)", prefix)
        operation = operation_match.group(1) if operation_match else ""
        concept = operation_match.group(2) if operation_match else prefix
        # Prefer the balance variation when available; otherwise preserve the PDF column position.
        previous_balance = state.get("santander_balance")
        is_debit = transaction_amount.start() < 58
        if previous_balance is not None and balance is not None and amount is not None:
            movement = abs(amount)
            delta = balance - previous_balance
            debit_error = abs(delta + movement)
            credit_error = abs(delta - movement)
            tolerance = max(0.02, movement * 0.001)
            if min(debit_error, credit_error) <= tolerance:
                is_debit = debit_error < credit_error
        last = {"Fecha": pending_date, "Operación": operation,
                "Concepto": re.sub(r"\s+", " ", concept).strip(),
                "Débito": abs(amount) if is_debit and amount is not None else None,
                "Crédito": abs(amount) if not is_debit and amount is not None else None,
                "Saldo": balance, "Origen": "", "Código trx": "", "Página": page_no,
                "Cuenta": state.get("account", "")}
        rows.append(last)
        state["santander_balance"] = balance
    return rows, rejected


def _parse_santander_plain_page(page: str, page_no: int, state: dict) -> tuple[list[dict], list[dict]]:
    """Fallback Santander parser independent from fixed character positions."""
    rows, rejected = [], []
    state["account"] = _account(page, "Santander", state.get("account", ""))
    active = False
    current_date = state.get("santander_date")
    previous_balance = state.get("santander_balance")
    for raw_line in page.splitlines():
        line = re.sub(r"\s+", " ", raw_line).strip()
        upper = line.upper()
        if "FECHA" in upper and "COMPROBANTE" in upper and "MOVIMIENTO" in upper and "SALDO" in upper:
            active = True
            continue
        if active and any(token in upper for token in ("DETALLE IMPOSITIVO", "SALDO TOTAL", "BANCO SANTANDER ARGENTINA S.A.")):
            active = False
        if not active or not line:
            continue
        date_match = re.match(r"^(\d{1,2}/\d{1,2}/\d{2,4})(?:\s+(.*))?$", line)
        remainder = line
        if date_match:
            current_date = _date(date_match.group(1))
            state["santander_date"] = current_date
            remainder = (date_match.group(2) or "").strip()
        amounts = list(SANTANDER_MONEY.finditer(line))
        if "SALDO INICIAL" in upper and amounts:
            previous_balance = ar_number(amounts[-1].group())
            if amounts[-1].group().lstrip().startswith("-"):
                previous_balance = -abs(previous_balance)
            state["santander_balance"] = previous_balance
            continue
        if current_date is None or len(amounts) < 2 or upper.startswith("RESPONSABLE:"):
            if rows and not amounts and remainder and not re.match(r"^\d+\s*-\s*\d+$", remainder):
                rows[-1]["Concepto"] = re.sub(r"\s+", " ", f'{rows[-1]["Concepto"]} {remainder}').strip()
            continue
        transaction_match, balance_match = amounts[-2], amounts[-1]
        transaction = abs(ar_number(transaction_match.group()) or 0.0)
        balance = ar_number(balance_match.group())
        if balance_match.group().lstrip().startswith("-"):
            balance = -abs(balance)
        prefix = remainder[:max(0, transaction_match.start() - (len(line) - len(remainder)))].strip()
        operation_match = re.match(r"(\d+)\s+(.*)", prefix)
        operation = operation_match.group(1) if operation_match else ""
        concept = operation_match.group(2) if operation_match else prefix
        if previous_balance is not None and balance is not None:
            is_credit = balance - previous_balance >= -0.01
        else:
            is_credit = any(token in upper for token in ("DEPOSITO", "ACREDIT", "TRANSFERENCIA RECIBIDA"))
        rows.append({"Fecha": current_date, "Operación": operation,
                     "Concepto": re.sub(r"\s+", " ", concept).strip(),
                     "Débito": None if is_credit else transaction, "Crédito": transaction if is_credit else None,
                     "Saldo": balance, "Origen": "", "Código trx": "", "Página": page_no,
                     "Cuenta": state.get("account", "")})
        previous_balance = balance
        state["santander_balance"] = balance
    return rows, rejected


def parse_pdf(pdf_bytes: bytes, forced_bank: str | None = None,
              progress: Callable[[int, int], None] | None = None):
    rows: list[dict] = []
    rejected: list[dict] = []
    state: dict = {}
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        total = len(pdf.pages)
        samples = []
        for sample_page in pdf.pages[:min(3, total)]:
            samples.append(sample_page.extract_text() or "")
            sample_page.close()
        bank = forced_bank if forced_bank and forced_bank != "Automático" else detect_bank("\n".join(samples))
        parsers = {"BTF": _parse_btf_page, "Macro": _parse_macro_page, "HSBC": _parse_hsbc_page,
                   "Comafi": _parse_comafi_page, "Santander": _parse_santander_page}
        for page_no, page in enumerate(pdf.pages, 1):
            if bank == "Macro":
                text = ""
                page_rows, page_rejected = _parse_macro_page_object(page, page_no, state)
            elif bank == "Comafi":
                text = ""
                page_rows, page_rejected = _parse_comafi_page_object(page, page_no, state)
            elif bank == "Galicia":
                text = page.extract_text() or ""
                page_rows, page_rejected = _parse_galicia_page(text, page_no, state)
            else:
                text = page.extract_text(layout=True, x_density=7.25, y_density=13) or ""
            if bank in {"Patagonia", "BBVA"}:
                page_rows, page_rejected = _parse_column_page(text, bank, page_no, state)
            elif bank not in {"Galicia", "Macro", "Comafi"} and bank in parsers:
                if bank == "BTF" and "LIQUIDACION DE PRESENTACION DE CUPONES" in text.upper():
                    page_rows, page_rejected = _parse_btf_liquidation_page(text, page_no, state)
                else:
                    page_rows, page_rejected = parsers[bank](text, page_no, state)
                if bank == "Santander" and not page_rows:
                    plain_text = page.extract_text() or ""
                    page_rows, page_rejected = _parse_santander_plain_page(plain_text, page_no, state)
            elif bank not in {"Galicia", "Macro", "Comafi"}:
                page_rows, page_rejected = [], []
            rows.extend(page_rows)
            rejected.extend(page_rejected)
            if state.get("account"):
                state.setdefault("accounts", set()).add(str(state["account"]).strip())
            if progress and (page_no == total or page_no % 5 == 0):
                progress(page_no, total)
            page.close()
            if page_no % 25 == 0:
                gc.collect()
    if bank == "Desconocido":
        rejected.append({"Página": None, "Texto": "", "Motivo": "Banco no reconocido"})
    base_columns = ["Banco", "Fecha", "Operación", "Concepto", "Débito", "Crédito",
                    "Saldo", "Origen", "Código trx", "Página", "Cuenta", "Mes"]
    frame = pd.DataFrame(rows)
    if not frame.empty:
        for column in ("Débito", "Crédito", "Saldo"):
            frame[column] = pd.to_numeric(frame[column], errors="coerce")
        if "Cuenta" not in frame.columns:
            frame["Cuenta"] = "N/D"
        frame["Cuenta"] = frame["Cuenta"].fillna("").astype(str).str.strip().replace(
            {"": "N/D", "nan": "N/D", "None": "N/D"}
        )
        frame.insert(0, "Banco", bank)
        frame["Mes"] = frame["Fecha"].dt.to_period("M").astype(str)
    else:
        # A recognized statement with no transactions is a valid normalization,
        # not a processing error.  Keep typed, stable columns so the application
        # can still generate an auditable Excel workbook.
        frame = pd.DataFrame({column: pd.Series(dtype="object") for column in base_columns})
        frame["Fecha"] = pd.to_datetime(frame["Fecha"])
        for column in ("Débito", "Crédito", "Saldo"):
            frame[column] = pd.to_numeric(frame[column], errors="coerce")
    frame.attrs["accounts"] = sorted(
        account for account in state.get("accounts", set()) if account
    )
    frame.attrs["status"] = (
        "SIN MOVIMIENTOS EN EL PERÍODO" if frame.empty and bank != "Desconocido" else "OK"
    )
    return bank, frame, pd.DataFrame(rejected)
