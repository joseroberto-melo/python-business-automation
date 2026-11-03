#!/usr/bin/env python3
"""Generate a synthetic payments workbook with deliberate data-quality problems.

All data is fictional: names are random combinations, e-mails use the reserved
``example.com`` domain and amounts are random numbers. Nothing here comes from
a real company or customer.

The workbook is seeded, so the same seed always produces the same file. The
function also returns a "manifest" with the exact number of problems injected,
which the tests use to check that the pipeline finds all of them.
"""
from __future__ import annotations

import argparse
import random
from datetime import date, timedelta
from itertools import product
from pathlib import Path

import pandas as pd

FIRST_NAMES = [
    "Ana", "Bruno", "Carla", "Diego", "Elena", "Felipe", "Gabriela", "Hugo",
    "Isabela", "Joaquin", "Karen", "Lucas", "Marta", "Nicolas", "Olivia",
    "Pablo", "Renata", "Sofia", "Tomas", "Valentina",
]
LAST_NAMES = [
    "Silva", "Santos", "Oliveira", "Gonzalez", "Rojas", "Martinez", "Costa",
    "Pereira", "Fernandez", "Lopes",
]
CURRENCIES = ["USD", "BRL", "EUR", "MXN", "GBP"]
CURRENCY_WEIGHTS = [0.40, 0.25, 0.15, 0.10, 0.10]
STATUSES = ["paid", "pending", "refunded"]
STATUS_WEIGHTS = [0.70, 0.20, 0.10]

BASE_ROWS = 120
N_CUSTOMERS = 40
COLUMNS = [
    "payment_id", "customer_id", "customer_name", "email",
    "amount", "currency", "payment_date", "status",
]

# How many rows get each problem. Each affected row gets exactly one problem,
# so the counts below are also the expected rejection counts.
INJECTED_ERRORS = {
    "invalid_email": 6,
    "non_positive_amount": 5,
    "invalid_amount": 4,
    "missing_customer_id": 4,
    "invalid_date": 4,
    "future_date": 2,
    "invalid_currency": 3,
    "invalid_status": 3,
}
MESSY_VALID_ROWS = 8   # valid rows with stray spaces / odd casing (should be cleaned)
DUPLICATE_ROWS = 6     # exact copies of valid rows (should be rejected)


def _make_customers(rng: random.Random) -> list[dict]:
    pairs = rng.sample(list(product(FIRST_NAMES, LAST_NAMES)), N_CUSTOMERS)
    return [
        {
            "customer_id": f"C{i:04d}",
            "customer_name": f"{first} {last}",
            "email": f"{first}.{last}@example.com".lower(),
        }
        for i, (first, last) in enumerate(pairs, start=1)
    ]


def _make_base_rows(rng: random.Random, customers: list[dict]) -> list[dict]:
    start = date(2026, 1, 5)
    span_days = (date(2026, 10, 28) - start).days
    rows = []
    for i in range(1, BASE_ROWS + 1):
        customer = rng.choice(customers)
        day = start + timedelta(days=rng.randint(0, span_days))
        # Real exports mix date formats; keep both ISO and DD/MM/YYYY.
        date_text = day.strftime("%Y-%m-%d") if rng.random() < 0.7 else day.strftime("%d/%m/%Y")
        rows.append(
            {
                "payment_id": f"P{i:05d}",
                **customer,
                "amount": round(rng.uniform(15, 2500), 2),
                "currency": rng.choices(CURRENCIES, CURRENCY_WEIGHTS)[0],
                "payment_date": date_text,
                "status": rng.choices(STATUSES, STATUS_WEIGHTS)[0],
            }
        )
    return rows


def _inject_error(rng: random.Random, row: dict, kind: str, n: int) -> None:
    if kind == "invalid_email":
        row["email"] = ["maria.example.com", "joao@", "@example.com", "pedro@@example.com",
                        "ana silva@example.com", "carlos@example"][n % 6]
    elif kind == "non_positive_amount":
        row["amount"] = [-50.0, 0, -12.34, -999.99, 0.0][n % 5]
    elif kind == "invalid_amount":
        row["amount"] = ["abc", "N/A", "12,50", ""][n % 4]
    elif kind == "missing_customer_id":
        row["customer_id"] = ""
    elif kind == "invalid_date":
        row["payment_date"] = ["31/02/2026", "2026-13-45", "yesterday", "14-03-2026"][n % 4]
    elif kind == "future_date":
        row["payment_date"] = ["2031-01-15", "2030-06-30"][n % 2]
    elif kind == "invalid_currency":
        row["currency"] = ["XXX", "BTC", "ARS"][n % 3]
    elif kind == "invalid_status":
        row["status"] = ["failed", "unknown", "chargeback"][n % 3]
    else:  # pragma: no cover - guards against typos in INJECTED_ERRORS
        raise ValueError(kind)


def _make_messy(row: dict) -> None:
    """Valid data that is only badly formatted: the pipeline should clean it."""
    row["email"] = f"  {row['email'].upper()} "
    row["currency"] = row["currency"].lower()
    row["status"] = f"{row['status'].upper()} "


def generate(path: str | Path, seed: int = 42) -> dict:
    """Write the workbook to ``path`` and return the manifest of injected problems."""
    rng = random.Random(seed)
    customers = _make_customers(rng)
    rows = _make_base_rows(rng, customers)

    n_errors = sum(INJECTED_ERRORS.values())
    chosen = rng.sample(range(BASE_ROWS), n_errors + MESSY_VALID_ROWS)
    error_idx, messy_idx = chosen[:n_errors], chosen[n_errors:]

    cursor = 0
    for kind, count in INJECTED_ERRORS.items():
        for n in range(count):
            _inject_error(rng, rows[error_idx[cursor]], kind, n)
            cursor += 1
    for idx in messy_idx:
        _make_messy(rows[idx])

    # Duplicates: copy clean rows and insert the copy somewhere after the original.
    touched = set(chosen)
    clean_idx = [i for i in range(BASE_ROWS) if i not in touched]
    # Snapshot the sources first: inserting shifts positions, so indices go stale.
    sources = [dict(rows[idx]) for idx in rng.sample(clean_idx, DUPLICATE_ROWS)]
    for source in sources:
        position = next(i for i, r in enumerate(rows) if r["payment_id"] == source["payment_id"])
        rows.insert(rng.randint(position + 1, len(rows)), dict(source))

    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    pd.DataFrame(rows, columns=COLUMNS).to_excel(path, index=False, sheet_name="payments")

    return {
        "total_rows": len(rows),
        "expected_valid": BASE_ROWS - n_errors,
        "expected_rejected": n_errors + DUPLICATE_ROWS,
        "expected_valid_rows_normalized": MESSY_VALID_ROWS,
        "expected_rejection_breakdown": {**INJECTED_ERRORS, "duplicate_payment_id": DUPLICATE_ROWS},
    }


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--output", default="data/input/payments_raw.xlsx")
    parser.add_argument("--seed", type=int, default=42)
    args = parser.parse_args()

    manifest = generate(args.output, args.seed)
    print(f"Wrote {manifest['total_rows']} synthetic rows to {args.output}")
    print(f"Injected problems: {manifest['expected_rejected']} rows should be rejected, "
          f"{manifest['expected_valid']} should pass.")


if __name__ == "__main__":
    main()
