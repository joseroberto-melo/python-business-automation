#!/usr/bin/env python3
"""Payments cleanup and enrichment pipeline.

Reads a raw payments workbook, validates every row, converts amounts to USD
using a public exchange-rate REST API, and writes:

* ``processed_payments.xlsx``  valid rows (+ USD amount) and a Summary sheet
* ``rejected_records.xlsx``    invalid rows with the reason(s) for rejection
* ``logs/run_<timestamp>.log`` a step-by-step log of the run

Every input row ends up in exactly one of the two workbooks.
"""
from __future__ import annotations

import argparse
import json
import logging
import math
import re
import sys
import time
from collections import Counter
from datetime import date, datetime
from pathlib import Path

import pandas as pd
import requests
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

LOGGER_NAME = "payments"
log = logging.getLogger(LOGGER_NAME)

REQUIRED_COLUMNS = [
    "payment_id", "customer_id", "customer_name", "email",
    "amount", "currency", "payment_date", "status",
]
ALLOWED_CURRENCIES = {"USD", "EUR", "BRL", "MXN", "GBP"}
ALLOWED_STATUSES = {"paid", "pending", "refunded"}
DATE_FORMATS = ("%Y-%m-%d", "%d/%m/%Y")
EMAIL_RE = re.compile(r"^[A-Za-z0-9._%+-]+@[A-Za-z0-9-]+(\.[A-Za-z0-9-]+)+$")

RATE_API_URL = "https://api.frankfurter.dev/v1/latest"
DEFAULT_FALLBACK_FILE = Path(__file__).parent / "data" / "reference" / "fallback_rates.json"


# --------------------------------------------------------------------------- #
# Validation
# --------------------------------------------------------------------------- #
def parse_amount(text: str) -> float | None:
    """Return the amount as a float, or None if it is not a plain number."""
    try:
        value = float(text.strip())
    except (ValueError, AttributeError):
        return None
    return value if math.isfinite(value) else None


def parse_date(text: str) -> date | None:
    """Accept ISO (2026-03-14) or DD/MM/YYYY (14/03/2026); None if invalid."""
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(text.strip(), fmt).date()
        except ValueError:
            continue
    return None


def normalize_row(raw: dict) -> dict:
    """Trim whitespace and standardise casing. Does not judge validity."""
    return {
        "payment_id": raw["payment_id"].strip(),
        "customer_id": raw["customer_id"].strip(),
        "customer_name": raw["customer_name"].strip(),
        "email": raw["email"].strip().lower(),
        "amount": raw["amount"].strip(),
        "currency": raw["currency"].strip().upper(),
        "payment_date": raw["payment_date"].strip(),
        "status": raw["status"].strip().lower(),
    }


def validate_row(row: dict, seen_ids: set[str], today: date) -> list[str]:
    """Return a list of reason codes. An empty list means the row is valid."""
    reasons: list[str] = []

    if not row["payment_id"]:
        reasons.append("missing_payment_id")
    elif row["payment_id"] in seen_ids:
        reasons.append("duplicate_payment_id")

    if not row["customer_id"]:
        reasons.append("missing_customer_id")

    if not EMAIL_RE.match(row["email"]):
        reasons.append("invalid_email")

    amount = parse_amount(row["amount"])
    if amount is None:
        reasons.append("invalid_amount")
    elif amount <= 0:
        reasons.append("non_positive_amount")

    if row["currency"] not in ALLOWED_CURRENCIES:
        reasons.append("invalid_currency")

    parsed_date = parse_date(row["payment_date"])
    if parsed_date is None:
        reasons.append("invalid_date")
    elif parsed_date > today:
        reasons.append("future_date")

    if row["status"] not in ALLOWED_STATUSES:
        reasons.append("invalid_status")

    return reasons


# --------------------------------------------------------------------------- #
# Enrichment: exchange rates over a REST API
# --------------------------------------------------------------------------- #
class RateProvider:
    """Returns the rate to convert one unit of a currency into USD.

    * One API call per distinct currency (results are cached for the run).
    * Failed calls are retried with exponential backoff.
    * If the API stays unreachable, static fallback rates are used and every
      affected row is labelled ``rate_source = "fallback"`` so nothing is hidden.
    """

    def __init__(self, offline: bool = False, session: requests.Session | None = None,
                 fallback_file: Path = DEFAULT_FALLBACK_FILE, max_attempts: int = 3,
                 backoff_seconds: float = 1.0, timeout: float = 10.0):
        self.offline = offline
        self.session = session or requests.Session()
        self.fallback_file = Path(fallback_file)
        self.max_attempts = max_attempts
        self.backoff_seconds = backoff_seconds
        self.timeout = timeout
        self._cache: dict[str, tuple[float, str]] = {}
        self._fallback: dict[str, float] | None = None

    def get_rate(self, currency: str) -> tuple[float, str]:
        if currency == "USD":
            return 1.0, "identity"
        if currency not in self._cache:
            self._cache[currency] = self._lookup(currency)
        return self._cache[currency]

    def _lookup(self, currency: str) -> tuple[float, str]:
        if self.offline:
            log.info("Offline mode: using static rate for %s", currency)
            return self._fallback_rate(currency), "offline"
        try:
            return self._fetch_from_api(currency), "api"
        except requests.RequestException as exc:
            log.warning("Rate API unavailable for %s (%s); using fallback rate", currency, exc)
            return self._fallback_rate(currency), "fallback"

    def _fetch_from_api(self, currency: str) -> float:
        params = {"base": currency, "symbols": "USD"}
        last_error: Exception | None = None
        for attempt in range(1, self.max_attempts + 1):
            try:
                response = self.session.get(RATE_API_URL, params=params, timeout=self.timeout)
                response.raise_for_status()
                rate = float(response.json()["rates"]["USD"])
                log.info("API rate %s->USD = %s (attempt %d)", currency, rate, attempt)
                return rate
            except (requests.RequestException, KeyError, ValueError) as exc:
                last_error = exc
                log.warning("Rate request for %s failed on attempt %d/%d: %s",
                            currency, attempt, self.max_attempts, type(exc).__name__)
                if attempt < self.max_attempts:
                    time.sleep(self.backoff_seconds * 2 ** (attempt - 1))
        # Bad payloads (KeyError/ValueError) are treated like a network failure.
        raise requests.RequestException(
            f"all {self.max_attempts} attempts failed ({type(last_error).__name__})")

    def _fallback_rate(self, currency: str) -> float:
        if self._fallback is None:
            self._fallback = json.loads(self.fallback_file.read_text(encoding="utf-8"))
        return float(self._fallback[currency])


# --------------------------------------------------------------------------- #
# Output helpers
# --------------------------------------------------------------------------- #
HEADER_FONT = Font(bold=True)
HEADER_FILL = PatternFill("solid", fgColor="D9E1F2")


def _style_sheet(ws, date_columns: tuple[str, ...] = ()) -> None:
    """Bold header, frozen first row, sensible column widths, date format."""
    for cell in ws[1]:
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
    ws.freeze_panes = "A2"
    headers = [c.value for c in ws[1]]
    for idx, column_cells in enumerate(ws.columns, start=1):
        longest = max(len(str(c.value)) if c.value is not None else 0 for c in column_cells)
        ws.column_dimensions[get_column_letter(idx)].width = min(longest + 2, 45)
        if headers[idx - 1] in date_columns:
            for cell in column_cells[1:]:
                cell.number_format = "yyyy-mm-dd"


def _write_summary_sheet(wb, summary: dict) -> None:
    ws = wb.create_sheet("Summary")
    rows: list[tuple] = [
        ("Run summary", None),
        ("Input file", summary["input_file"]),
        ("Rows read", summary["input_rows"]),
        ("Valid rows", summary["valid_rows"]),
        ("Rejected rows", summary["rejected_rows"]),
        ("Valid rows cleaned (spaces/casing)", summary["valid_rows_normalized"]),
        ("Total valid amount (USD)", summary["total_amount_usd"]),
        (None, None),
        ("Rejections by reason", "Rows"),
        *summary["rejection_breakdown"].items(),
        (None, None),
        ("Valid rows by currency", "Rows"),
        *[(f"{cur} ({info['rate_source']}, rate {info['rate']})", info["rows"])
          for cur, info in summary["by_currency"].items()],
    ]
    for row in rows:
        ws.append(row)
    for header_row in (1, 9, 9 + len(summary["rejection_breakdown"]) + 2):
        for cell in ws[header_row]:
            cell.font = HEADER_FONT
            cell.fill = HEADER_FILL
    ws.column_dimensions["A"].width = 42
    ws.column_dimensions["B"].width = 30


# --------------------------------------------------------------------------- #
# Pipeline
# --------------------------------------------------------------------------- #
def run(input_path: str | Path, output_dir: str | Path = "output", log_dir: str | Path = "logs",
        offline: bool = False, today: date | None = None,
        provider: RateProvider | None = None) -> dict:
    """Run the full pipeline and return a summary dict."""
    input_path, output_dir, log_dir = Path(input_path), Path(output_dir), Path(log_dir)
    today = today or date.today()
    output_dir.mkdir(parents=True, exist_ok=True)
    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = log_dir / f"run_{datetime.now():%Y%m%d_%H%M%S}.log"
    handlers = _setup_logging(log_file)
    try:
        return _run(input_path, output_dir, log_file, offline, today, provider)
    finally:
        for handler in handlers:
            log.removeHandler(handler)
            handler.close()


def _setup_logging(log_file: Path) -> list[logging.Handler]:
    formatter = logging.Formatter("%(asctime)s | %(levelname)-7s | %(message)s")
    handlers = [logging.FileHandler(log_file, encoding="utf-8"), logging.StreamHandler(sys.stdout)]
    log.setLevel(logging.INFO)
    for handler in handlers:
        handler.setFormatter(formatter)
        log.addHandler(handler)
    return handlers


def _run(input_path: Path, output_dir: Path, log_file: Path, offline: bool, today: date,
         provider: RateProvider | None) -> dict:
    log.info("Run started | input=%s | offline=%s | reference date=%s", input_path, offline, today)

    # Step 1 - load. Everything is read as text so nothing is silently coerced.
    log.info("Step 1/5 Loading input")
    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")
    raw = pd.read_excel(input_path, dtype=str, keep_default_na=False)
    missing = [c for c in REQUIRED_COLUMNS if c not in raw.columns]
    if missing:
        raise ValueError(f"Input is missing required columns: {missing}")
    log.info("Loaded %d rows", len(raw))

    # Steps 2+3 - normalise and validate each row.
    log.info("Step 2/5 Normalising and validating rows")
    valid: list[dict] = []
    rejected: list[dict] = []
    seen_ids: set[str] = set()
    normalized_valid = 0
    for position, record in enumerate(raw[REQUIRED_COLUMNS].to_dict("records")):
        excel_row = position + 2  # header is row 1
        row = normalize_row(record)
        reasons = validate_row(row, seen_ids, today)
        if row["payment_id"]:
            seen_ids.add(row["payment_id"])
        if reasons:
            rejected.append({"source_row": excel_row, **record, "reasons": "; ".join(reasons)})
            log.warning("Row %d (%s) rejected: %s", excel_row, row["payment_id"] or "no id", "; ".join(reasons))
        else:
            if any(record[k] != row[k] for k in REQUIRED_COLUMNS):
                normalized_valid += 1
                log.info("Row %d (%s) cleaned (whitespace/casing)", excel_row, row["payment_id"])
            row["amount"] = round(float(row["amount"]), 2)
            row["payment_date"] = parse_date(row["payment_date"])
            valid.append(row)
    log.info("Validation done: %d valid, %d rejected", len(valid), len(rejected))

    # Step 4 - enrich with USD amounts.
    log.info("Step 3/5 Enriching valid rows with USD amounts")
    provider = provider or RateProvider(offline=offline)
    rates_used: dict[str, tuple[float, str]] = {}
    for row in valid:
        rate, source = provider.get_rate(row["currency"])
        rates_used[row["currency"]] = (rate, source)
        row["usd_rate"] = rate
        row["amount_usd"] = round(row["amount"] * rate, 2)
        row["rate_source"] = source

    # Step 5 - write outputs.
    log.info("Step 4/5 Building summary")
    processed_df = pd.DataFrame(valid, columns=REQUIRED_COLUMNS + ["usd_rate", "amount_usd", "rate_source"])
    rejected_df = pd.DataFrame(rejected, columns=["source_row", *REQUIRED_COLUMNS, "reasons"])

    breakdown = Counter(code for r in rejected for code in r["reasons"].split("; "))
    by_currency = {
        cur: {"rows": int((processed_df["currency"] == cur).sum()), "rate": rate, "rate_source": src}
        for cur, (rate, src) in sorted(rates_used.items())
    }
    processed_path = output_dir / "processed_payments.xlsx"
    rejected_path = output_dir / "rejected_records.xlsx"
    summary = {
        "input_file": str(input_path),
        "input_rows": len(raw),
        "valid_rows": len(valid),
        "rejected_rows": len(rejected),
        "valid_rows_normalized": normalized_valid,
        "total_amount_usd": round(float(processed_df["amount_usd"].sum()), 2),
        "rejection_breakdown": dict(sorted(breakdown.items())),
        "by_currency": by_currency,
        "processed_file": str(processed_path),
        "rejected_file": str(rejected_path),
        "log_file": str(log_file),
    }

    log.info("Step 5/5 Writing output workbooks")
    with pd.ExcelWriter(processed_path, engine="openpyxl") as writer:
        processed_df.to_excel(writer, index=False, sheet_name="Processed")
        _style_sheet(writer.sheets["Processed"], date_columns=("payment_date",))
        _write_summary_sheet(writer.book, summary)
    with pd.ExcelWriter(rejected_path, engine="openpyxl") as writer:
        rejected_df.to_excel(writer, index=False, sheet_name="Rejected")
        _style_sheet(writer.sheets["Rejected"])

    # Every input row must be accounted for in exactly one output.
    assert len(valid) + len(rejected) == len(raw), "row count mismatch"
    log.info("Run finished | valid=%d rejected=%d total_usd=%.2f", len(valid), len(rejected),
             summary["total_amount_usd"])
    return summary


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Validate, enrich and report on a payments workbook.")
    parser.add_argument("--input", default="data/input/payments_raw.xlsx")
    parser.add_argument("--output-dir", default="output")
    parser.add_argument("--log-dir", default="logs")
    parser.add_argument("--offline", action="store_true",
                        help="Skip the exchange-rate API and use data/reference/fallback_rates.json")
    args = parser.parse_args(argv)

    try:
        summary = run(args.input, args.output_dir, args.log_dir, offline=args.offline)
    except (FileNotFoundError, ValueError) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2

    print("\n=== Summary ===")
    print(f"Rows read:     {summary['input_rows']}")
    print(f"Valid:         {summary['valid_rows']}")
    print(f"Rejected:      {summary['rejected_rows']}")
    print(f"Total (USD):   {summary['total_amount_usd']:,.2f}")
    print(f"Outputs:       {summary['processed_file']}, {summary['rejected_file']}")
    print(f"Log:           {summary['log_file']}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
