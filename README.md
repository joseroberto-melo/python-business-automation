# Python Business Automation: Payments Cleanup & Enrichment

A small, runnable Python pipeline that takes a messy Excel export of customer payments, validates every row, converts amounts to USD through a public REST API, and produces clean reports plus an audit log.

It is a typical back-office automation job: the kind of manual spreadsheet work that finance and operations teams repeat every week.

> **All data in this repository is synthetic.** Names are random combinations, e-mails use the reserved `example.com` domain, and amounts are random. Nothing comes from a real company or customer.

## The problem

Payment exports from different systems arrive with the same kinds of defects:

- invalid or missing e-mails, customer IDs and payment IDs
- negative, zero or non-numeric amounts
- mixed date formats and impossible dates
- unsupported currencies and unknown statuses
- duplicated rows
- stray spaces and inconsistent casing

Fixing this by hand is slow and error-prone, and nobody can later prove what was changed or rejected.

## What the pipeline does

```
payments_raw.xlsx
      │
      ▼
 1. Load (everything read as text, nothing silently coerced)
 2. Normalise (trim spaces, standardise casing)
 3. Validate (collect every reason a row fails, not just the first)
 4. Enrich (convert amount to USD via REST API: retries, caching, fallback)
 5. Report (processed workbook + rejected workbook + run log)
      │
      ▼
output/processed_payments.xlsx   valid rows + USD amount + Summary sheet
output/rejected_records.xlsx     invalid rows + source row number + reasons
logs/run_<timestamp>.log         step-by-step audit trail
```

Every input row ends up in exactly one of the two workbooks, and the run asserts this before finishing.

## Validation rules

| Reason code | Rule |
|---|---|
| `missing_payment_id` / `duplicate_payment_id` | `payment_id` must be present and unique (first occurrence wins) |
| `missing_customer_id` | `customer_id` must be present |
| `invalid_email` | must look like `name@domain.tld` |
| `invalid_amount` | must be a plain number (`12,50` and `N/A` are rejected on purpose, since the decimal separator is ambiguous) |
| `non_positive_amount` | must be greater than zero |
| `invalid_currency` | one of `USD, EUR, BRL, MXN, GBP` |
| `invalid_date` / `future_date` | `YYYY-MM-DD` or `DD/MM/YYYY`, a real calendar date, not in the future |
| `invalid_status` | one of `paid, pending, refunded` |

A row can fail several rules at once, and all reasons are listed in `rejected_records.xlsx`.

## Quick start

Requires Python 3.10+.

```bash
git clone <this-repo-url>
cd python-business-automation

python -m venv .venv
# Windows:      .venv\Scripts\activate
# macOS/Linux:  source .venv/bin/activate

pip install -r requirements.txt

python generate_sample_data.py     # creates data/input/payments_raw.xlsx (126 rows, seeded)
python process_payments.py         # runs the pipeline with live exchange rates
```

No API key is needed. To run without internet access:

```bash
python process_payments.py --offline
```

Run the tests:

```bash
pytest
```

## Example result

The sample workbook has 126 rows with deliberately injected problems. A run produces:

| | Rows |
|---|---|
| Rows read | 126 |
| Valid (written to `processed_payments.xlsx`) | 89 |
| Rejected (written to `rejected_records.xlsx`) | 37 |
| of which duplicates | 6 |

A committed sample run is in [`examples/`](examples/). It was generated with `--offline`, so its exchange rates are the static illustrative ones from `data/reference/fallback_rates.json`, and the `rate_source` column says so.

## Design notes

- **Exchange rates over REST.** Rates come from the free [Frankfurter](https://frankfurter.dev) API (ECB reference rates, no key). There is one call per distinct currency, cached for the run.
- **Retries with exponential backoff.** Each call is attempted up to 3 times (1 s, 2 s between attempts), with a timeout.
- **Graceful degradation, never silent.** If the API stays unreachable, the run continues with static fallback rates, and every affected row is labelled `rate_source = fallback` and logged as a warning. The Summary sheet shows which source each currency used.
- **Nothing is dropped quietly.** Rejected rows keep their original values and Excel row number, so a person can fix them at the source.
- **Reads as text.** Excel and pandas love to guess types; reading everything as text keeps validation in the pipeline's hands.
- **Tested.** The test suite covers each validation rule, retry and fallback behaviour (with a fake HTTP session), and an end-to-end run that checks the pipeline finds exactly the problems the generator injected.

## Known simplifications

This is a demonstration, so a few things are intentionally kept simple:

- Amounts are converted at the **latest** rate, not the rate on each payment date. A production version would use historical rates.
- The input is one Excel file with a fixed schema. Real jobs usually add configuration for column names and rules.
- Scheduling, e-mail notifications and database/ERP output are not included, but they are natural extensions.

## Project structure

```
generate_sample_data.py     synthetic data generator (seeded, with a manifest of injected problems)
process_payments.py         the pipeline (validation, enrichment, reporting, logging, CLI)
data/reference/             static fallback exchange rates
tests/test_pipeline.py      unit and end-to-end tests
examples/                   a sample run (offline)
requirements.txt
```

## About

Built by José Roberto Melo, a Python and RPA developer focused on business process automation (Python, UiPath, Power Automate, Automation Anywhere, REST/SOAP integrations, OCR).
