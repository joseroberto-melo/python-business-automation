"""Tests for the payments pipeline. Run with:  pytest"""
from datetime import date

import pandas as pd
import pytest
import requests

import generate_sample_data
import process_payments as pp

TODAY = date(2026, 11, 3)


class FakeResponse:
    def __init__(self, payload=None, status=200):
        self._payload, self.status_code = payload, status

    def raise_for_status(self):
        if self.status_code >= 400:
            raise requests.HTTPError(f"HTTP {self.status_code}")

    def json(self):
        return self._payload


class FakeSession:
    """Returns the queued responses/exceptions in order and records the calls."""

    def __init__(self, *outcomes):
        self.outcomes, self.calls = list(outcomes), []

    def get(self, url, params=None, timeout=None):
        self.calls.append(params)
        outcome = self.outcomes.pop(0)
        if isinstance(outcome, Exception):
            raise outcome
        return outcome


# ----------------------------- row-level rules ----------------------------- #
def make_row(**overrides):
    row = {"payment_id": "P1", "customer_id": "C1", "customer_name": "Ana Silva",
           "email": "ana@example.com", "amount": "10.50", "currency": "USD",
           "payment_date": "2026-03-14", "status": "paid"}
    row.update(overrides)
    return row


@pytest.mark.parametrize("overrides, expected", [
    ({}, []),
    ({"email": "ana.example.com"}, ["invalid_email"]),
    ({"amount": "-5"}, ["non_positive_amount"]),
    ({"amount": "12,50"}, ["invalid_amount"]),
    ({"customer_id": ""}, ["missing_customer_id"]),
    ({"payment_date": "31/02/2026"}, ["invalid_date"]),
    ({"payment_date": "2031-01-01"}, ["future_date"]),
    ({"currency": "BTC"}, ["invalid_currency"]),
    ({"status": "failed"}, ["invalid_status"]),
    ({"email": "bad", "amount": "0"}, ["invalid_email", "non_positive_amount"]),
])
def test_validate_row(overrides, expected):
    row = pp.normalize_row(make_row(**overrides))
    assert pp.validate_row(row, set(), TODAY) == expected


def test_duplicate_payment_id_is_flagged():
    row = pp.normalize_row(make_row())
    assert pp.validate_row(row, {"P1"}, TODAY) == ["duplicate_payment_id"]


def test_both_date_formats_are_accepted():
    assert pp.parse_date("2026-03-14") == pp.parse_date("14/03/2026") == date(2026, 3, 14)


def test_normalisation_cleans_spaces_and_case():
    row = pp.normalize_row(make_row(email="  ANA@EXAMPLE.COM ", currency="usd", status="PAID "))
    assert (row["email"], row["currency"], row["status"]) == ("ana@example.com", "USD", "paid")


# --------------------------------- rates ----------------------------------- #
def provider(session, **kwargs):
    return pp.RateProvider(session=session, backoff_seconds=0, **kwargs)


def test_usd_needs_no_api_call():
    session = FakeSession()
    assert provider(session).get_rate("USD") == (1.0, "identity")
    assert session.calls == []


def test_rate_is_fetched_once_and_cached():
    session = FakeSession(FakeResponse({"rates": {"USD": 1.1}}))
    rates = provider(session)
    assert rates.get_rate("EUR") == (1.1, "api")
    assert rates.get_rate("EUR") == (1.1, "api")
    assert len(session.calls) == 1


def test_retries_then_succeeds():
    session = FakeSession(requests.ConnectionError("boom"), FakeResponse(status=503),
                          FakeResponse({"rates": {"USD": 0.2}}))
    assert provider(session).get_rate("BRL") == (0.2, "api")
    assert len(session.calls) == 3


def test_falls_back_after_all_attempts_fail():
    session = FakeSession(*[requests.Timeout("slow")] * 3)
    rate, source = provider(session).get_rate("BRL")
    assert source == "fallback" and rate == 0.18
    assert len(session.calls) == 3


def test_malformed_api_payload_triggers_fallback():
    session = FakeSession(*[FakeResponse({"unexpected": True})] * 3)
    assert provider(session).get_rate("MXN")[1] == "fallback"


# ------------------------------ end to end --------------------------------- #
@pytest.fixture
def pipeline_result(tmp_path):
    source = tmp_path / "payments_raw.xlsx"
    manifest = generate_sample_data.generate(source, seed=42)
    summary = pp.run(source, tmp_path / "out", tmp_path / "logs", offline=True, today=TODAY)
    return manifest, summary, tmp_path


def test_every_injected_problem_is_found(pipeline_result):
    manifest, summary, _ = pipeline_result
    assert summary["input_rows"] == manifest["total_rows"]
    assert summary["valid_rows"] == manifest["expected_valid"]
    assert summary["rejected_rows"] == manifest["expected_rejected"]
    assert summary["rejection_breakdown"] == dict(sorted(manifest["expected_rejection_breakdown"].items()))
    assert summary["valid_rows_normalized"] == manifest["expected_valid_rows_normalized"]


def test_outputs_are_consistent(pipeline_result):
    manifest, summary, tmp_path = pipeline_result
    processed = pd.read_excel(tmp_path / "out" / "processed_payments.xlsx", sheet_name="Processed")
    rejected = pd.read_excel(tmp_path / "out" / "rejected_records.xlsx")
    assert len(processed) + len(rejected) == manifest["total_rows"]
    assert processed["payment_id"].is_unique
    assert (processed["amount_usd"] > 0).all()
    assert set(processed["rate_source"]) <= {"identity", "offline"}
    assert rejected["reasons"].notna().all()
    assert abs(processed["amount_usd"].sum() - summary["total_amount_usd"]) < 0.01
    assert list((tmp_path / "logs").glob("run_*.log"))


def test_missing_columns_raise_clear_error(tmp_path):
    bad = tmp_path / "bad.xlsx"
    pd.DataFrame({"payment_id": ["P1"]}).to_excel(bad, index=False)
    with pytest.raises(ValueError, match="missing required columns"):
        pp.run(bad, tmp_path / "out", tmp_path / "logs", offline=True)


def test_missing_input_file_raises(tmp_path):
    with pytest.raises(FileNotFoundError):
        pp.run(tmp_path / "nope.xlsx", tmp_path / "out", tmp_path / "logs", offline=True)
