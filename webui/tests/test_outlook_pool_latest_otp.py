from __future__ import annotations

import email
from datetime import datetime, timezone

from webui.backend import outlook_pool


def _msg(raw: str):
    return email.message_from_string(raw)


def test_message_body_text_extracts_html_before_otp_regex():
    msg = _msg(
        "Content-Type: text/html; charset=utf-8\n"
        "Subject: ChatGPT code\n"
        "\n"
        "<html><body><span>Your verification code is</span>"
        "<strong>123456</strong></body></html>"
    )

    text = outlook_pool._message_body_text(msg)

    assert "Your verification code is" in text
    assert outlook_pool._extract_otp_from_html(text) == "123456"


def test_extract_otp_ignores_hex_color_context():
    text = "style color:#353740; Your code is 778899"

    assert outlook_pool._extract_otp_from_html(text) == "778899"


class FakeImap:
    def __init__(self, messages_by_folder):
        self.messages_by_folder = messages_by_folder
        self.current_folder = ""

    def authenticate(self, *_args, **_kwargs):
        return "OK", []

    def list(self):
        return "OK", [b'(\\HasNoChildren) "/" "INBOX"', b'(\\HasNoChildren) "/" "Junk"']

    def select(self, folder, readonly=True):
        self.current_folder = folder.strip('"')
        if self.current_folder not in self.messages_by_folder:
            return "NO", []
        return "OK", []

    def uid(self, command, _arg, *args):
        if command == "SEARCH":
            count = len(self.messages_by_folder[self.current_folder])
            return "OK", [b" ".join(str(i).encode() for i in range(1, count + 1))]
        if command == "FETCH":
            uid = int(_arg)
            return "OK", [(b"1 (BODY[] {1}", self.messages_by_folder[self.current_folder][uid - 1])]
        raise AssertionError(command)

    def logout(self):
        return "OK", []


def _raw_mail(sender: str, subject: str, body: str, ts: int) -> bytes:
    date = datetime.fromtimestamp(ts, timezone.utc).strftime("%a, %d %b %Y %H:%M:%S +0000")
    return (
        f"From: {sender}\n"
        f"Subject: {subject}\n"
        f"Date: {date}\n"
        "Content-Type: text/html; charset=utf-8\n"
        "\n"
        f"{body}"
    ).encode()


def test_fetch_otp_uses_latest_message_per_folder(monkeypatch):
    messages = {
        "INBOX": [
            _raw_mail("noreply@tm.openai.com", "Old code", "Your code is 111111", 1000),
            _raw_mail("noreply@tm.openai.com", "New code", "Your code is 222222", 2000),
        ],
        "Junk": [
            _raw_mail("newsletter@example.com", "Latest non otp", "Invoice 999999", 3000),
        ],
    }
    monkeypatch.setattr(outlook_pool, "get_outlook_access_token", lambda *_: "access-token")
    monkeypatch.setattr(outlook_pool, "fetch_otp_via_graph", lambda *_args, **_kwargs: "")
    monkeypatch.setattr(outlook_pool.imaplib, "IMAP4_SSL", lambda *_: FakeImap(messages))

    otp = outlook_pool.fetch_otp_via_imap(
        "owner@outlook.com", "refresh", "client", timeout=60, threshold_ts=1
    )

    assert otp == "222222"


def test_fetch_otp_via_graph_extracts_latest_openai_message(monkeypatch):
    responses = {
        "inbox": {
            "value": [{
                "subject": "Welcome",
                "from": {"emailAddress": {"name": "Microsoft", "address": "microsoft@example.com"}},
                "receivedDateTime": "2026-05-21T10:00:00Z",
                "bodyPreview": "Welcome",
                "body": {"content": "Welcome"},
            }]
        },
        "junkemail": {
            "value": [{
                "subject": "Your ChatGPT code",
                "from": {"emailAddress": {"name": "OpenAI", "address": "noreply@tm.openai.com"}},
                "receivedDateTime": "2026-05-21T10:01:00Z",
                "bodyPreview": "Your verification code is 456789",
                "body": {"content": "<b>Your verification code is 456789</b>"},
            }]
        },
    }

    class FakeResp:
        def __init__(self, payload):
            self.payload = payload

        def __enter__(self):
            return self

        def __exit__(self, *_args):
            return False

        def read(self):
            import json
            return json.dumps(self.payload).encode()

    def fake_urlopen(req, timeout=20):
        url = req.full_url
        folder = "junkemail" if "junkemail" in url else "inbox"
        return FakeResp(responses[folder])

    monkeypatch.setattr(outlook_pool, "get_outlook_access_token", lambda *_args, **_kwargs: "access-token")
    monkeypatch.setattr(outlook_pool.urllib.request, "urlopen", fake_urlopen)

    otp = outlook_pool.fetch_otp_via_graph("owner@outlook.com", "refresh", "client")

    assert otp == "456789"
