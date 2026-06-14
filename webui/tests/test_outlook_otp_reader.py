from __future__ import annotations

import email
import importlib.util
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]
SPEC = importlib.util.spec_from_file_location(
    "outlook_otp_reader",
    ROOT / "scripts" / "outlook_otp_reader.py",
)
reader = importlib.util.module_from_spec(SPEC)
sys.modules[SPEC.name] = reader
SPEC.loader.exec_module(reader)  # type: ignore[union-attr]


def test_parse_account_line_accepts_four_part_supplier_format():
    account = reader.parse_account_line(
        "Owner@outlook.com----secret----client-123----refresh-token"
    )

    assert account.email == "owner@outlook.com"
    assert account.password == "secret"
    assert account.client_id == "client-123"
    assert account.refresh_token == "refresh-token"


def test_parse_account_line_rejects_bad_format():
    try:
        reader.parse_account_line("owner@outlook.com----only-two")
    except reader.OutlookOtpError as exc:
        assert "email----password----client_id----refresh_token" in str(exc)
    else:
        raise AssertionError("bad account line should fail")


def test_extract_otp_prefers_keyword_context_over_plain_number():
    text = "Invoice 987654 was created. Your verification code is 112233."

    assert reader.extract_otp(text) == "112233"


def test_extract_otp_handles_code_before_keyword():
    text = "Use 445566 as your login verification code."

    assert reader.extract_otp(text) == "445566"


def test_extract_otp_ignores_hex_color_context():
    text = "style color:#123456; Your code is 778899"

    assert reader.extract_otp(text) == "778899"


def test_body_from_message_extracts_html_text():
    msg = email.message_from_string(
        "Content-Type: text/html; charset=utf-8\n"
        "Subject: Verify\n"
        "\n"
        "<html><body><p>Your code is <b>123456</b></p></body></html>"
    )

    body = reader._body_from_message(msg)

    assert "Your code is" in body
    assert "123456" in body


def test_message_matches_requires_all_configured_filters():
    assert reader.message_matches(
        sender="Example <security@example.com>",
        subject="Your login code",
        body="Use 123456 to continue",
        from_contains=["example.com"],
        subject_contains=["login"],
        body_contains=["continue"],
    )
    assert not reader.message_matches(
        sender="Example <security@example.com>",
        subject="Your login code",
        body="Use 123456 to continue",
        from_contains=["example.com"],
        subject_contains=["billing"],
        body_contains=["continue"],
    )
