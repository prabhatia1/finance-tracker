"""
PDF Statement Parser — Extracts transactions from password-protected bank/credit card PDFs.

Supports:
  - HDFC (bank account + credit card)
  - SBI (bank account + credit card)
  - ICICI (bank account + credit card)
  - Axis / generic layouts

Parsing strategies, in order of preference:
  1. Coordinate engine  — reconstructs rows from word x/y positions (handles HDFC
                          net-banking PDFs, which have no ruling lines between rows).
  2. Table engine       — pdfplumber extract_tables() for PDFs with real ruled tables.
  3. Text/regex engine  — last-resort line scanning.

Usage:
    from pdf_parser import parse_pdf
    transactions, errors = parse_pdf(pdf_path_or_fileobj, password="1234")
"""

import re
from datetime import datetime

try:
    import pdfplumber
    _HAS_PDFPLUMBER = True
except ImportError:
    pdfplumber = None
    _HAS_PDFPLUMBER = False


# ─── Amount parsing ───────────────────────────────────────────────────────────

def _clean_amount(val):
    """Parse currency: '1,234.56', 'Rs. 1,234.56', '(1,234.56)', '1,234.56 DR', '1,234.56 C'."""
    if val is None:
        return 0.0
    s = str(val).strip()
    if not s:
        return 0.0

    # Parentheses = negative (some banks show credits as (1,234.56))
    is_negative = False
    if s.startswith("(") and s.endswith(")"):
        is_negative = True
        s = s[1:-1].strip()

    # Trailing DR/CR indicators (HDFC/ICICI use Dr/Cr, SBI cards use D/C)
    is_credit_suffix = False
    is_debit_suffix = False
    upper = s.upper()
    if upper.endswith("DR") or upper.endswith("DBT"):
        is_debit_suffix = True
        s = s[:-2].strip()
    elif upper.endswith("CR") or upper.endswith("CRT"):
        is_credit_suffix = True
        s = s[:-2].strip()
    elif re.search(r'\d\s*D$', upper):
        is_debit_suffix = True
        s = s[:-1].strip()
    elif re.search(r'\d\s*C$', upper):
        is_credit_suffix = True
        s = s[:-1].strip()

    s = s.replace("Rs.", "").replace("Rs", "").replace("INR", "").replace("₹", "")
    s = s.replace(",", "").replace(" ", "").replace("\xa0", "")

    try:
        amt = float(s)
    except ValueError:
        return 0.0

    if is_negative or is_credit_suffix:
        return -abs(amt)  # negative = credit
    if is_debit_suffix:
        return abs(amt)
    return amt


# ─── Date parsing ─────────────────────────────────────────────────────────────

_DATE_FORMATS = [
    "%d/%m/%Y", "%d-%m-%Y", "%d/%m/%y", "%d-%m-%y",
    "%Y-%m-%d", "%d %b %Y", "%d-%b-%Y", "%d %B %Y",
    "%d-%b-%y", "%d %b %y", "%b %d, %Y", "%B %d, %Y",
    "%d %b, %Y", "%d-%B-%Y", "%d.%m.%Y", "%d.%m.%y",
    "%d%b%Y", "%d%b%y",
]


def _parse_date(raw_date, formats=None):
    """Parse a date string. Returns ISO format (YYYY-MM-DD) or None."""
    if not raw_date:
        return None
    raw_date = str(raw_date).strip().replace("\xa0", " ")
    if not raw_date:
        return None

    for fmt in (formats or _DATE_FORMATS):
        try:
            return datetime.strptime(raw_date, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue

    # Regex: pull out date-like substring
    m = re.search(r'(\d{1,2}[/.-]\d{1,2}[/.-]\d{2,4})', raw_date)
    if m:
        for fmt in ["%d/%m/%Y", "%d-%m-%Y", "%d/%m/%y", "%d-%m-%y", "%d.%m.%Y", "%d.%m.%y"]:
            try:
                return datetime.strptime(m.group(1), fmt).strftime("%Y-%m-%d")
            except ValueError:
                continue
    m = re.search(r'(\d{1,2}[\s-]+\w{3,9}[\s-]+\d{2,4})', raw_date)
    if m:
        cleaned = re.sub(r'[\s-]+', ' ', m.group(1))
        for fmt in ["%d %b %Y", "%d %B %Y", "%d %b %y", "%d %B %y"]:
            try:
                return datetime.strptime(cleaned, fmt).strftime("%Y-%m-%d")
            except ValueError:
                continue
    return None


# ─── Token classification (shared by all engines) ─────────────────────────────

# An amount must have a decimal part or thousands separators, so reference
# numbers like 0000621760665968 are never mistaken for money.
_AMOUNT_TOKEN_RE = re.compile(r'^\(?(?:Rs\.?|INR|₹)?\s*'
                              r'(?:\d{1,3}(?:,\d{2,3})+(?:\.\d{1,2})?|\d+\.\d{1,2})'
                              r'\)?\s*(?:DR|CR|DBT|CRT|D|C)?$', re.IGNORECASE)
# Inside a known numeric column a bare integer is acceptable too.
_LOOSE_AMOUNT_TOKEN_RE = re.compile(r'^\(?(?:Rs\.?|INR|₹)?\s*'
                                    r'\d{1,3}(?:,\d{2,3})*(?:\.\d{1,2})?'
                                    r'\)?\s*(?:DR|CR|DBT|CRT|D|C)?$', re.IGNORECASE)
_DATE_TOKEN_RE = re.compile(r'^\d{1,2}[/.-]\d{1,2}[/.-]\d{2,4}$')
_DATE_WORD_RE = re.compile(r'^\d{1,2}[\s-]?[A-Za-z]{3,9}[\s-]?\d{2,4}$')
_TYPE_TOKENS = {"dr", "cr", "d", "c", "debit", "credit", "dbt", "crt"}

_AMOUNT_RE = re.compile(r'\(?[\d,]+\.\d{1,2}\)?')
_DATE_RE = re.compile(r'\d{1,2}[/.-]\d{1,2}[/.-]\d{2,4}')


def _is_amount_token(text, loose=False):
    if not text:
        return False
    t = str(text).strip()
    if not t:
        return False
    return bool((_LOOSE_AMOUNT_TOKEN_RE if loose else _AMOUNT_TOKEN_RE).match(t))


def _is_date_token(text):
    if not text:
        return False
    t = str(text).strip()
    return bool(_DATE_TOKEN_RE.match(t) or _DATE_WORD_RE.match(t))


def _compact(text):
    """Lowercase, letters+digits only — used for fuzzy header/keyword matching."""
    return re.sub(r'[^a-z0-9]', '', str(text or "").lower())


def _extract_amounts(text):
    """Extract all amount values from a string."""
    if not text:
        return []
    out = []
    for m in _AMOUNT_RE.findall(str(text)):
        v = _clean_amount(m)
        if v != 0:
            out.append(abs(v))
    return out


def _extract_dates(text):
    """Extract all date values from a string."""
    if not text:
        return []
    dates = []
    for m in _DATE_RE.findall(str(text)):
        d = _parse_date(m)
        if d:
            dates.append(d)
    return dates


# ─── Header field classification ──────────────────────────────────────────────

def _classify_header(text):
    """Map a header cell to a logical field name, or None if unrecognised."""
    k = _compact(text)
    if not k:
        return None

    # Value date must be checked before the generic date rule (HDFC has both
    # "Date" and "ValueDt" — only the first is the transaction date).
    if k.startswith("value") and ("dt" in k or "date" in k):
        return "value_date"
    if k.startswith("post") and ("dt" in k or "date" in k):
        return "value_date"

    if "withdrawal" in k or "moneyout" in k or "amountwithdrawn" in k:
        return "debit"
    if k in ("debit", "dr", "debits", "dramount", "debited") or k.startswith("debitamt") \
            or k.startswith("debitamount") or k.startswith("debitrs") or k.startswith("debitinr"):
        return "debit"
    if "deposit" in k or "moneyin" in k or "amountdeposited" in k:
        return "credit"
    if k in ("credit", "cr", "credits", "cramount", "credited") or k.startswith("creditamt") \
            or k.startswith("creditamount") or k.startswith("creditrs") or k.startswith("creditinr"):
        return "credit"
    if "balance" in k or k in ("bal", "closingbal", "openingbal"):
        return "balance"
    if k in ("drcr", "crdr", "type", "txntype", "transactiontype", "dc", "cd",
             "debitcredit", "creditdebit", "trxntype"):
        return "type_col"
    if "chq" in k or "cheque" in k or k == "utr" or "refno" in k or k.startswith("ref") \
            or "referenceno" in k or "refnumber" in k:
        return "ref"
    if "serno" in k or k in ("srno", "sno", "slno", "serialno"):
        return "serial"
    if "rewardpoint" in k or k == "points":
        return "points"
    # ICICI cards put an "Intl.Amount" column before the real rupee amount;
    # claiming it as the amount column would lose every transaction value.
    if k.startswith("intl") or "internationalamount" in k or "forexamount" in k:
        return "intl_amount"
    if "date" in k or k in ("dt", "txndt", "trandt", "trxndate"):
        return "date"
    if any(x in k for x in ("narration", "narrative", "particular", "description",
                            "detail", "remark", "merchant")):
        return "desc"
    if "amount" in k or k in ("amt", "value", "txnamt", "billamt"):
        return "amount"
    if k in ("transaction", "transactions", "txn"):
        return "desc"
    return None


# Legacy header sets (used by the table engine) ────────────────────────────────
DATE_HEADERS = {
    "date", "txn date", "transaction date", "posting date", "value date",
    "date of transaction", "txn dt", "trans date", "book date",
    "date dd/mm/yyyy", "trxn date", "date dd/mm/yy", "transaction dt", "post date",
}
DESC_HEADERS = {
    "description", "narration", "narrative", "particulars", "transaction details",
    "remarks", "details", "merchant", "merchant name", "transaction description",
    "transaction", "txn description", "description of transaction",
    "nature of transaction", "transaction remarks", "name of the merchant",
}
AMOUNT_HEADERS = {
    "amount", "txn amount", "transaction amount", "amount inr", "bill amount",
    "total amount", "value", "amount rs", "amount in rs",
}
DEBIT_HEADERS = {
    "debit", "dr", "debit amount", "withdrawal", "withdrawals", "debit rs",
    "withdrawal rs", "debit inr", "dr amount", "withdrawal amount",
    "amount withdrawn", "debited", "money out",
}
CREDIT_HEADERS = {
    "credit", "cr", "credit amount", "deposit", "deposits", "credit rs",
    "deposit rs", "credit inr", "cr amount", "deposit amount",
    "amount deposited", "credited", "money in",
}
TYPE_INDICATOR_HEADERS = {
    "dr cr", "cr dr", "type", "txn type", "transaction type", "d c", "c d",
    "debit credit", "credit debit", "dr/cr", "cr/dr", "d/c", "c/d",
}
BALANCE_HEADERS = {
    "balance", "closing balance", "running balance", "bal", "balance rs",
    "balance inr", "book balance", "avail balance", "ledger balance", "closing bal",
}
REF_HEADERS = {
    "cheque no", "cheque number", "ref no", "ref number", "reference no",
    "reference number", "cheque/ref no", "chq no", "chq number",
    "transaction ref", "txn ref", "utr", "reference",
}


def _normalize_header(h):
    """Normalize a header string for matching."""
    if not h:
        return ""
    h = str(h).lower().strip()
    h = re.sub(r'[().\-_/]', ' ', h)
    return re.sub(r'\s+', ' ', h).strip()


def _map_columns(headers):
    """Map column indices for a ruled table. Returns dict or None."""
    mapping = {"date": None, "desc": None, "amount": None, "debit": None,
               "credit": None, "type_col": None, "balance": None, "ref": None}
    for i, h in enumerate(headers):
        field = _classify_header(h)
        if field in mapping and mapping[field] is None:
            mapping[field] = i

    if mapping["date"] is None:
        for i, h in enumerate(headers):
            if "date" in _normalize_header(h):
                mapping["date"] = i
                break
    if mapping["desc"] is None:
        for i, h in enumerate(headers):
            nh = _normalize_header(h)
            if any(k in nh for k in ["desc", "narr", "particular", "remark", "detail", "merchant"]):
                mapping["desc"] = i
                break

    has_amount = any(mapping[k] is not None for k in ("amount", "debit", "credit"))
    if mapping["date"] is None or not has_amount:
        return None
    return mapping


# ─── Noise / summary filtering ────────────────────────────────────────────────

SKIP_KEYWORDS = [
    "opening balance", "closing balance", "opening bal", "closing bal",
    "total debit", "total credit", "grand total", "sub total", "subtotal",
    "summary", "continuation", "continued", "brought forward", "b/f",
    "carried forward", "c/f", "page total", "page subtotal",
    "net total", "period total", "statement total", "total transactions",
    "available balance", "ledger balance", "balance carried",
    "balance brought", "total amount", "total debits", "total credits",
    "account summary", "transaction summary", "card summary",
]

# Compact (letters+digits only) fragments that mark statement chrome, not data.
_NOISE_FRAGMENTS = [
    "statementsummary", "pageno", "generatedon", "generatedby",
    "requestingbranch", "computergenerated", "doesnotrequiresignature",
    "registeredoffice", "contentsofthisstatement", "closingbalanceincludes",
    "accountbranch", "accountstatus", "accounttype", "accountno", "custid",
    "jointholders", "nomination", "odlimit", "statementofaccount",
    "openingbalance", "drcount", "crcount", "branchcode",
    "rtgsneftifsc", "gstinnumber", "hdfcbanklimited", "unbilledtransactions",
    "pleasenote", "importantinformation", "rewardpointssummary",
    "minimumamountdue", "paymentduedate", "totalamountdue", "creditlimit",
    "availablecreditlimit", "availablecashlimit", "statementperiod",
    "transactiondetails", "endofstatement", "thisisasystemgenerated",
]

_STOP_FRAGMENTS = ["statementsummary", "endofstatement"]


def _is_summary_row(desc):
    """Check if a description is a summary/balance row that should be skipped."""
    lower = str(desc or "").lower().strip()
    return any(k in lower for k in SKIP_KEYWORDS)


def _is_noise_line(text):
    c = _compact(text)
    if not c:
        return True
    return any(f in c for f in _NOISE_FRAGMENTS)


def _is_header_row(row_cells):
    """Check if a row looks like a table header."""
    joined = " ".join(str(c or "").lower() for c in row_cells)
    has_date = any(k in joined for k in ["date", "txn dt", "post date"])
    has_amount = any(k in joined for k in
                     ["amount", "debit", "credit", "withdrawal", "deposit", "narration"])
    return has_date and has_amount


CREDIT_KEYWORDS = [
    "salary", "deposit", "refund", "cashback", "reversal", "return",
    "received", "interest", "dividend", "cash deposit", "cheque deposit",
    "clearing credit", "inward", "money received", "payment received",
    "repayment", "acct credit", "mmt credit", "trf cr", "trf credit",
    "chqdep", "clg", "credit adjustment", "reward redemption",
]


def _looks_like_credit(desc):
    lower = str(desc or "").lower()
    return any(k in lower for k in CREDIT_KEYWORDS)


def _new_txn(date, desc, amount, txn_type):
    return {
        "date": date,
        "description": str(desc or "")[:200],
        "amount": round(abs(amount), 2),
        "category": None,
        "card_id": None,
        "txn_type": txn_type,
        "source": "upload",
    }


# ─── Coordinate engine ────────────────────────────────────────────────────────
# HDFC net-banking statements draw one big box around the transaction area with
# no lines between rows, so extract_tables() collapses every transaction into a
# single cell. Rebuilding rows from word coordinates avoids that entirely.

_NUMERIC_FIELDS = ("debit", "credit", "amount", "balance")


def _page_lines(page, y_tol=3.0):
    """Group a page's words into visual lines sorted top-to-bottom."""
    try:
        words = page.extract_words(keep_blank_chars=False, use_text_flow=False)
    except Exception:
        return []

    lines = []
    for w in sorted(words, key=lambda w: (round(w["top"], 1), w["x0"])):
        for ln in lines:
            if abs(ln["top"] - w["top"]) <= y_tol:
                ln["words"].append(w)
                break
        else:
            lines.append({"top": w["top"], "words": [w]})

    for ln in lines:
        ln["words"].sort(key=lambda w: w["x0"])
        ln["text"] = " ".join(w["text"] for w in ln["words"])
    lines.sort(key=lambda l: l["top"])
    return lines


def _merge_header_cells(words, gap=4.0):
    """Merge adjacent header words ('Withdrawal' + 'Amt.') into one cell."""
    cells = []
    for w in sorted(words, key=lambda w: w["x0"]):
        if cells and (w["x0"] - cells[-1]["x1"]) <= gap:
            cells[-1]["text"] += " " + w["text"]
            cells[-1]["x1"] = max(cells[-1]["x1"], w["x1"])
        else:
            cells.append({"text": w["text"], "x0": w["x0"], "x1": w["x1"]})
    return cells


def _detect_columns(line):
    """If this line is a transaction-table header, return its column layout."""
    cells = _merge_header_cells(line["words"])
    if len(cells) < 3:
        return None

    columns = []
    seen = set()
    for c in cells:
        field = _classify_header(c["text"])
        # Keep only the first occurrence of each field; unknown cells are kept
        # as anonymous anchors so column boundaries stay correct.
        if field is None or field in seen:
            field = None
        else:
            seen.add(field)
        columns.append({"field": field, "x0": c["x0"], "x1": c["x1"], "text": c["text"]})

    if "date" not in seen:
        return None
    if not any(f in seen for f in ("debit", "credit", "amount")):
        return None
    if "desc" not in seen:
        return None

    # Right boundary of a column = left edge of the next column's header.
    columns.sort(key=lambda c: c["x0"])
    for i, c in enumerate(columns):
        c["limit"] = columns[i + 1]["x0"] if i + 1 < len(columns) else float("inf")
    return columns


def _assign_to_column(word, columns):
    """Assign a word to a column using its right edge (numbers are right-aligned)."""
    x1 = word["x1"]
    for c in columns:
        if x1 <= c["limit"] + 0.5:
            return c
    return columns[-1]


def _blank_row():
    return {"date": None, "desc": [], "ref": None, "debit": None, "credit": None,
            "amount": None, "balance": None, "type": None, "extra_amounts": []}


def _row_from_line(line, columns):
    """Slot a line's words into columns. Returns a row dict, or None if empty."""
    row = _blank_row()
    got_anything = False
    date_tokens = []

    for w in line["words"]:
        col = _assign_to_column(w, columns)
        field = col["field"]
        text = w["text"].strip()
        if not text:
            continue
        got_anything = True

        if field == "date":
            # SBI cards print "05 Aug 25" as three separate words, so buffer the
            # whole column and parse it once the line is consumed.
            date_tokens.append(text)
            continue

        if field in _NUMERIC_FIELDS:
            if _is_amount_token(text, loose=True):
                val = abs(_clean_amount(text))
                signed = _clean_amount(text)
                if row[field] is None:
                    row[field] = val
                else:
                    row["extra_amounts"].append((field, val))
                if signed < 0 and field == "amount":
                    row["type"] = "credit"
                continue
            if _compact(text) in _TYPE_TOKENS:
                row["type"] = "credit" if _compact(text).startswith("c") else "debit"
                continue
            # Text that overflowed into a money column still belongs to the narration.
            row["desc"].append(text)
            continue

        if field == "type_col":
            k = _compact(text)
            if k in _TYPE_TOKENS:
                row["type"] = "credit" if k.startswith("c") else "debit"
            continue

        if field == "ref":
            if row["ref"] is None:
                row["ref"] = text
            continue

        if field == "desc":
            row["desc"].append(text)
            continue

        # value_date / serial / points / intl_amount / unknown anchors are ignored.

    if date_tokens:
        # Try the longest prefix that parses, so "05 Aug 25 FOO" still works.
        for n in range(len(date_tokens), 0, -1):
            parsed = _parse_date(" ".join(date_tokens[:n]))
            if parsed:
                row["date"] = parsed
                # Whatever followed the date is narration that overflowed left.
                row["desc"] = date_tokens[n:] + row["desc"]
                break
        else:
            # Nothing date-like here: this is a continuation line whose text
            # started left of the description column.
            row["desc"] = date_tokens + row["desc"]

    return row if got_anything else None


def _headerless_row(line):
    """Parse a line without a known column layout: date … description … amounts."""
    words = line["words"]
    if not words:
        return None

    row = _blank_row()
    idx = 0

    # Leading date, possibly split across tokens ("05 Aug 2026").
    for span in (3, 2, 1):
        if len(words) >= span:
            candidate = " ".join(w["text"] for w in words[:span])
            if _is_date_token(candidate) or _parse_date(candidate):
                parsed = _parse_date(candidate)
                if parsed:
                    row["date"] = parsed
                    idx = span
                    break
    # Trailing amounts (plus an optional Dr/Cr marker after them).
    tail = []
    j = len(words) - 1
    while j >= idx:
        text = words[j]["text"].strip()
        k = _compact(text)
        if k in _TYPE_TOKENS and not tail:
            row["type"] = "credit" if k.startswith("c") else "debit"
            j -= 1
            continue
        if _is_amount_token(text):
            tail.insert(0, abs(_clean_amount(text)))
            if _clean_amount(text) < 0:
                row["type"] = "credit"
            j -= 1
            continue
        break

    row["desc"] = [w["text"] for w in words[idx:j + 1]]
    row["extra_amounts"] = [("amount", v) for v in tail[1:]] if tail else []
    if tail:
        row["amount"] = tail[0]
        if len(tail) > 1:
            # Two trailing numbers on a bank statement line is usually
            # "amount, running balance".
            row["amount"] = tail[0]
            row["balance"] = tail[-1]
            row["extra_amounts"] = [("amount", v) for v in tail[1:-1]]
    return row


def _collect_rows(pdf):
    """Walk every page and return (rows, columns_found, warnings)."""
    rows = []
    warnings = []
    columns = None
    saw_header = False
    table_top = None  # y of the header row, reused to locate data on later pages

    for page_no, page in enumerate(pdf.pages, 1):
        lines = _page_lines(page)
        in_table = columns is not None
        # Pages after the header page repeat the account-info block, so start
        # each of them inside a "preamble" that must be closed explicitly.
        preamble = columns is not None

        for line in lines:
            text = line["text"]
            compact = _compact(text)

            if any(f in compact for f in _STOP_FRAGMENTS):
                in_table = False
                continue

            detected = _detect_columns(line)
            if detected:
                columns = detected
                saw_header = True
                in_table = True
                preamble = False
                table_top = line["top"]
                continue

            if preamble:
                # HDFC ends the repeated block with "Statement of account"; for
                # other layouts fall back to the header row's vertical position.
                if "statementofaccount" in compact:
                    preamble = False
                elif table_top is not None and line["top"] >= table_top - 4:
                    preamble = False
                    # This line is real data — fall through and parse it.
                else:
                    continue
                if "statementofaccount" in compact:
                    continue

            if not in_table:
                continue

            row = _row_from_line(line, columns) if columns else _headerless_row(line)
            if not row:
                continue

            has_money = any(row[f] is not None for f in ("debit", "credit", "amount"))
            if row["date"] and has_money:
                # A dated row carrying money is a transaction, full stop. Never
                # noise-filter it — real narrations contain words like "MICR"
                # that also appear in the statement's header block.
                pass
            elif _is_noise_line(text):
                continue
            elif row["date"] is None and not row["desc"]:
                continue

            row["page"] = page_no
            rows.append(row)

    return rows, saw_header, warnings


def _find_opening_balance(pdf):
    """Locate the statement's opening balance (used to type the first row)."""
    # HDFC only prints the statement summary on the last page — scan backwards.
    for page in reversed(pdf.pages):
        text = page.extract_text() or ""
        lines = [l.strip() for l in text.split("\n")]
        for i, line in enumerate(lines):
            c = _compact(line)
            if "openingbalance" not in c and not ("opening" in c and "bal" in c):
                continue
            amts = _extract_amounts(line)
            if amts:
                return amts[0]
            # HDFC prints the summary labels on one line and values on the next.
            for nxt in lines[i + 1:i + 3]:
                amts = _extract_amounts(nxt)
                if amts:
                    return amts[0]
    return None


def _rows_to_transactions(rows, opening_balance=None):
    """Convert column-mapped rows into transactions, reconciling with balances."""
    txns = []
    warnings = []
    last_txn = None  # most recent transaction, for continuation lines

    for row in rows:
        desc = " ".join(row["desc"]).strip()

        if row["date"] is None:
            # Continuation line — extend the previous narration.
            if last_txn is not None and desc and not _is_summary_row(desc):
                last_txn["description"] = (last_txn["description"] + " " + desc).strip()[:200]
            continue

        if _is_summary_row(desc):
            continue

        debit = row["debit"]
        credit = row["credit"]
        amount = row["amount"]
        balance = row["balance"]

        value = None
        txn_type = None
        if debit:
            value, txn_type = debit, "debit"
        if credit:
            value, txn_type = credit, "credit"
        if value is None and amount:
            value = amount
            txn_type = row["type"] or ("credit" if _looks_like_credit(desc) else "debit")
        if txn_type is None:
            txn_type = row["type"] or "debit"
        if row["type"] and not (debit or credit):
            txn_type = row["type"]

        if not value:
            continue

        txn = _new_txn(row["date"], desc, value, txn_type)
        # Kept only for the post-processing passes below, then stripped.
        txn["_balance"] = balance
        ref = (row["ref"] or "").strip()
        txn["_ref"] = ref if ref and ref.strip("0") and ref not in ("N/A", "-") else ""
        txns.append(txn)
        last_txn = txn

    # ── Balance reconciliation: the running balance is the ground truth for
    # debit vs credit, which also repairs any column-assignment mistakes. ──
    prev = opening_balance
    fixed = 0
    for txn in txns:
        bal = txn.get("_balance")
        if bal is None:
            prev = None
            continue
        if prev is not None:
            diff = round(bal - prev, 2)
            if abs(abs(diff) - txn["amount"]) <= 0.02 and abs(diff) > 0:
                expected = "credit" if diff > 0 else "debit"
                if txn["txn_type"] != expected:
                    txn["txn_type"] = expected
                    fixed += 1
            elif abs(diff) > 0.02 and txn["amount"] == 0:
                txn["amount"] = abs(diff)
                txn["txn_type"] = "credit" if diff > 0 else "debit"
        prev = bal

    if fixed:
        warnings.append(f"Corrected debit/credit on {fixed} row(s) using the running balance")

    # Append the reference number after the full (multi-line) narration.
    for txn in txns:
        txn.pop("_balance", None)
        ref = txn.pop("_ref", "")
        if ref:
            txn["description"] = (f"{txn['description']} [Ref: {ref}]".strip())[:200]

    return [t for t in txns if t["amount"] > 0], warnings


def _parse_coordinates(pdf):
    """Primary strategy: rebuild rows from word positions."""
    rows, saw_header, warnings = _collect_rows(pdf)
    if not rows:
        return [], warnings
    opening = _find_opening_balance(pdf)
    txns, warn2 = _rows_to_transactions(rows, opening_balance=opening)
    return txns, warnings + warn2


# ─── Table engine (ruled tables) ──────────────────────────────────────────────

def _extract_tables_from_pdf(pdf):
    """Extract all tables from all pages."""
    tables = []
    for page_num, page in enumerate(pdf.pages, 1):
        try:
            page_tables = page.extract_tables()
        except Exception:
            continue
        for t in page_tables:
            if t and len(t) > 1:
                tables.append((page_num, t))
    return tables


def _find_header_and_mapping(table):
    """Find the header row and return (header_idx, mapping)."""
    for i, row in enumerate(table[:5]):
        if row and _is_header_row(row):
            mapping = _map_columns(row)
            if mapping:
                return i, mapping
    if table:
        mapping = _map_columns(table[0])
        if mapping:
            return 0, mapping
    return None, None


def _parse_table_rows(table, header_idx, mapping, prev_balance=None):
    """Parse data rows of a ruled table. Handles continuation rows."""
    transactions = []
    errors = []
    last_txn = None
    running_balance = prev_balance

    for i, row in enumerate(table[header_idx + 1:], start=header_idx + 2):
        if not row:
            continue
        try:
            cells = [str(c).strip() if c else "" for c in row]

            def cell(field):
                idx = mapping.get(field)
                if idx is None or idx >= len(cells):
                    return ""
                return cells[idx]

            raw_date = cell("date")
            parsed_date = _parse_date(raw_date) if raw_date else None

            desc_idx = mapping["desc"] if mapping["desc"] is not None else (mapping["date"] + 1)
            raw_desc = cells[desc_idx].replace("\n", " ").strip() if desc_idx < len(cells) else ""

            if not parsed_date:
                if raw_desc and last_txn and not _is_summary_row(raw_desc) and not _is_header_row(row):
                    last_txn["description"] = (last_txn["description"] + " " + raw_desc)[:200]
                continue
            if not raw_desc or _is_summary_row(raw_desc) or _is_header_row(row):
                continue

            ref_val = cell("ref").strip()
            if ref_val and ref_val.strip("0") and ref_val not in ("N/A", "-"):
                raw_desc = f"{raw_desc} [Ref: {ref_val}]"

            amount = 0.0
            txn_type = "debit"

            debit_amt = abs(_clean_amount(cell("debit")))
            if debit_amt > 0:
                amount, txn_type = debit_amt, "debit"
            credit_amt = abs(_clean_amount(cell("credit")))
            if credit_amt > 0:
                amount, txn_type = credit_amt, "credit"

            type_val = _compact(cell("type_col"))
            if type_val in _TYPE_TOKENS:
                txn_type = "credit" if type_val.startswith("c") else "debit"

            raw_amt = cell("amount")
            if raw_amt and amount == 0:
                signed = _clean_amount(raw_amt)
                amount = abs(signed)
                if signed < 0:
                    txn_type = "credit"
                elif mapping["type_col"] is None and mapping["debit"] is None \
                        and mapping["credit"] is None and _looks_like_credit(raw_desc):
                    txn_type = "credit"

            bal_amts = _extract_amounts(cell("balance"))
            if bal_amts and running_balance is not None:
                diff = round(bal_amts[-1] - running_balance, 2)
                if abs(abs(diff) - amount) <= 0.02 and abs(diff) > 0:
                    txn_type = "credit" if diff > 0 else "debit"
                elif amount == 0 and abs(diff) > 0:
                    amount, txn_type = abs(diff), ("credit" if diff > 0 else "debit")
            if bal_amts:
                running_balance = bal_amts[-1]

            if amount == 0:
                continue

            txn = _new_txn(parsed_date, raw_desc, amount, txn_type)
            transactions.append(txn)
            last_txn = txn

        except Exception as e:
            errors.append(f"Row {i}: {e}")

    return transactions, errors


def _parse_tables(pdf, opening_balance=None):
    best, best_errors = [], []
    for _page_num, table in _extract_tables_from_pdf(pdf):
        header_idx, mapping = _find_header_and_mapping(table)
        if not mapping:
            continue
        txns, errs = _parse_table_rows(table, header_idx, mapping, prev_balance=opening_balance)
        if len(txns) > len(best):
            best, best_errors = txns, errs
    return best, best_errors


# ─── Text/regex engine (last resort) ──────────────────────────────────────────

TEXT_PATTERNS = [
    re.compile(r'(\d{1,2}[/.-]\d{1,2}[/.-]\d{2,4})\s+(.+?)\s+'
               r'([\d,]+\.\d{1,2}\s*(?:DR|CR|DBT|CRT|D|C)?)$', re.IGNORECASE),
    re.compile(r'(\d{1,2}[\s-]\w{3,9}[\s-]\d{2,4})\s+(.+?)\s+'
               r'([\d,]+\.\d{1,2}\s*(?:DR|CR|DBT|CRT|D|C)?)$', re.IGNORECASE),
    re.compile(r'(\d{1,2}[/.-]\d{1,2}[/.-]\d{2,4})\s+(.+?)\s+\(([\d,]+\.\d{1,2})\)$',
               re.IGNORECASE),
]


def _parse_text_fallback(pdf):
    """Extract transactions from raw text using regex (when everything else fails)."""
    transactions = []
    errors = []

    for page in pdf.pages:
        text = page.extract_text() or ""
        for line in text.split("\n"):
            line = line.strip()
            if not line or _is_summary_row(line) or _is_noise_line(line):
                continue
            for pattern in TEXT_PATTERNS:
                m = pattern.search(line)
                if not m:
                    continue
                parsed_date = _parse_date(m.group(1))
                if not parsed_date:
                    continue
                desc = m.group(2).strip()
                if _is_summary_row(desc):
                    continue
                signed = _clean_amount(m.group(3))
                if signed == 0:
                    continue
                txn_type = "credit" if signed < 0 else "debit"
                if signed > 0 and _looks_like_credit(desc):
                    txn_type = "credit"
                transactions.append(_new_txn(parsed_date, desc, abs(signed), txn_type))
                break

    return transactions, errors


# ─── Main entry point ─────────────────────────────────────────────────────────

def parse_pdf(pdf_source, password=""):
    """
    Parse a PDF bank/credit card statement.

    Args:
        pdf_source: file path (str/Path) or file-like object
        password: PDF password (string)

    Returns:
        (transactions, errors) where:
          transactions: list of dicts with date, description, amount, txn_type, source
          errors: list of error/warning strings
    """
    if not _HAS_PDFPLUMBER:
        return [], ["pdfplumber not installed — run: pip install pdfplumber"]

    try:
        pdf = pdfplumber.open(pdf_source, password=password or "")
    except Exception as e:
        msg = str(e)
        if "password" in msg.lower() or "encrypt" in msg.lower():
            return [], [f"Could not open PDF — wrong or missing password ({msg})"]
        return [], [f"Could not open PDF: {msg}"]

    errors = []
    transactions = []

    try:
        # ── Strategy 1: coordinate engine ──
        try:
            transactions, warns = _parse_coordinates(pdf)
            errors.extend(warns)
        except Exception as e:
            errors.append(f"Coordinate engine failed: {e}")

        if transactions:
            errors.insert(0, f"Extracted via layout engine ({len(transactions)} transactions)")
        else:
            # ── Strategy 2: ruled tables ──
            try:
                opening = _find_opening_balance(pdf)
                transactions, errs = _parse_tables(pdf, opening_balance=opening)
                if transactions:
                    errors.extend(errs)
                    errors.insert(0, f"Extracted from tables ({len(transactions)} transactions)")
            except Exception as e:
                errors.append(f"Table engine failed: {e}")

        if not transactions:
            # ── Strategy 3: text/regex ──
            try:
                transactions, errs = _parse_text_fallback(pdf)
                if transactions:
                    errors.extend(errs)
                    errors.insert(0, f"Extracted via text/regex ({len(transactions)} transactions)")
            except Exception as e:
                errors.append(f"Text engine failed: {e}")

        if not transactions:
            errors.insert(0, "No transactions found — the PDF format may not be supported. "
                            "Use Preview to see the detected layout.")

    except Exception as e:
        errors.append(f"Parse error: {e}")
    finally:
        pdf.close()

    return transactions, errors


def debug_pdf(pdf_source, password="", max_rows=12):
    """
    Debug helper: shows the detected column layout, the first parsed rows and a
    raw text sample. Useful when a specific bank's statement won't parse.
    """
    if not _HAS_PDFPLUMBER:
        return {"error": "pdfplumber not installed"}

    try:
        pdf = pdfplumber.open(pdf_source, password=password or "")
    except Exception as e:
        return {"error": f"Could not open PDF: {e}"}

    result = {"tables": [], "text_sample": "", "page_count": len(pdf.pages)}

    try:
        if pdf.pages:
            result["text_sample"] = (pdf.pages[0].extract_text() or "")[:2000]

        # Detected layout from the coordinate engine.
        for page_num, page in enumerate(pdf.pages[:3], 1):
            for line in _page_lines(page):
                cols = _detect_columns(line)
                if not cols:
                    continue
                rows, _saw, _w = _collect_rows(pdf)
                result["tables"].append({
                    "page": page_num,
                    "header_idx": 0,
                    "mapping": {c["field"]: round(c["x0"], 1) for c in cols if c["field"]},
                    "raw_rows": [{
                        "row": i,
                        "cells": [
                            r["date"] or "",
                            " ".join(r["desc"])[:60],
                            r["ref"] or "",
                            f"D:{r['debit']}" if r["debit"] else "",
                            f"C:{r['credit']}" if r["credit"] else "",
                            f"A:{r['amount']}" if r["amount"] else "",
                            f"BAL:{r['balance']}" if r["balance"] else "",
                        ],
                    } for i, r in enumerate(rows[:max_rows])],
                    "total_rows": len(rows),
                })
                break
            if result["tables"]:
                break

        # Ruled tables, if any.
        for page_num, table in _extract_tables_from_pdf(pdf)[:2]:
            header_idx, mapping = _find_header_and_mapping(table)
            result["tables"].append({
                "page": page_num,
                "header_idx": header_idx,
                "mapping": mapping,
                "raw_rows": [{"row": i, "cells": [str(c)[:50] if c else "" for c in row]}
                             for i, row in enumerate(table[:max_rows])],
                "total_rows": len(table),
            })

        result["opening_balance"] = _find_opening_balance(pdf)
    except Exception as e:
        result["error"] = str(e)
    finally:
        pdf.close()

    return result
