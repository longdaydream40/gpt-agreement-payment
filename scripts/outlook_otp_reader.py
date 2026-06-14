#!/usr/bin/env python3
"""Read OTP codes from an authorized Outlook mailbox.

This is intentionally a single-mailbox, read-only utility. It expects the
mailbox owner to provide OAuth credentials through environment variables or
CLI arguments and requires a sender/subject/body filter by default.

Environment variables:
    OUTLOOK_EMAIL
    OUTLOOK_CLIENT_ID
    OUTLOOK_REFRESH_TOKEN

Example:
    python scripts/outlook_otp_reader.py \
        --from-contains example.com \
        --subject-contains verify \
        --timeout 180
"""

from __future__ import annotations

import argparse
import dataclasses
import email
import email.header
import email.utils
import html
import imaplib
import json
import os
import re
import socket
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from dataclasses import dataclass
from html.parser import HTMLParser
from typing import Iterable, Optional, Sequence


TOKEN_URL = "https://login.microsoftonline.com/common/oauth2/v2.0/token"
IMAP_HOST = "outlook.office365.com"
IMAP_PORT = 993
IMAP_SCOPE = "https://outlook.office.com/IMAP.AccessAsUser.All offline_access"
DEFAULT_FOLDERS = ("INBOX", "Junk", "Junk Email", "Spam")


class OutlookOtpError(RuntimeError):
    """Base error for Outlook OTP reader failures."""


@dataclass(frozen=True)
class OutlookAccount:
    email: str
    client_id: str
    refresh_token: str
    password: str = ""


@dataclass(frozen=True)
class MailHit:
    folder: str
    uid: str
    sender: str
    subject: str
    date: str
    timestamp: float
    code: str


class _HtmlToText(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.parts: list[str] = []

    def handle_data(self, data: str) -> None:
        if data:
            self.parts.append(data)

    def text(self) -> str:
        return html.unescape(" ".join(self.parts))


def parse_account_line(value: str) -> OutlookAccount:
    """Parse email----password----client_id----refresh_token format."""
    parts = [part.strip() for part in (value or "").strip().split("----")]
    if len(parts) != 4:
        raise OutlookOtpError("account line must be: email----password----client_id----refresh_token")
    mailbox, password, client_id, refresh_token = parts
    if "@" not in mailbox:
        raise OutlookOtpError("account line email is invalid")
    if not client_id:
        raise OutlookOtpError("account line client_id is empty")
    if not refresh_token:
        raise OutlookOtpError("account line refresh_token is empty")
    return OutlookAccount(
        email=mailbox.lower(),
        password=password,
        client_id=client_id,
        refresh_token=refresh_token,
    )


def account_from_env_or_args(args: argparse.Namespace) -> OutlookAccount:
    if args.account_file:
        with open(args.account_file, "r", encoding="utf-8") as fh:
            for raw in fh:
                line = raw.strip()
                if line and not line.startswith("#"):
                    return parse_account_line(line)
        raise OutlookOtpError(f"no account line found in {args.account_file}")

    email_value = args.email or os.environ.get("OUTLOOK_EMAIL", "")
    client_id = args.client_id or os.environ.get("OUTLOOK_CLIENT_ID", "")
    refresh_token = args.refresh_token or os.environ.get("OUTLOOK_REFRESH_TOKEN", "")
    missing = [
        name
        for name, value in (
            ("OUTLOOK_EMAIL/--email", email_value),
            ("OUTLOOK_CLIENT_ID/--client-id", client_id),
            ("OUTLOOK_REFRESH_TOKEN/--refresh-token", refresh_token),
        )
        if not value
    ]
    if missing:
        raise OutlookOtpError("missing required credentials: " + ", ".join(missing))
    return OutlookAccount(
        email=email_value.strip().lower(),
        client_id=client_id.strip(),
        refresh_token=refresh_token.strip(),
    )


def get_access_token(account: OutlookAccount, timeout: int = 20) -> str:
    body = urllib.parse.urlencode(
        {
            "grant_type": "refresh_token",
            "refresh_token": account.refresh_token,
            "client_id": account.client_id,
            "scope": IMAP_SCOPE,
        }
    ).encode("utf-8")
    req = urllib.request.Request(
        TOKEN_URL,
        data=body,
        headers={"Content-Type": "application/x-www-form-urlencoded"},
    )
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            data = json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        raw = exc.read().decode("utf-8", errors="replace")
        try:
            data = json.loads(raw)
            detail = data.get("error_description") or data.get("error") or raw[:200]
        except Exception:
            detail = raw[:200]
        raise OutlookOtpError(f"token refresh failed: HTTP {exc.code}: {detail}") from exc
    except Exception as exc:
        raise OutlookOtpError(f"token refresh failed: {exc}") from exc

    token = data.get("access_token")
    if not token:
        raise OutlookOtpError(f"token response did not include access_token: {sorted(data.keys())}")
    return str(token)


def imap_login(account: OutlookAccount, access_token: str, socket_timeout: int = 20) -> imaplib.IMAP4_SSL:
    conn = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT)
    try:
        conn.socket().settimeout(socket_timeout)
    except (AttributeError, socket.error):
        pass
    auth = f"user={account.email}\x01auth=Bearer {access_token}\x01\x01"
    try:
        status, _ = conn.authenticate("XOAUTH2", lambda _: auth.encode("utf-8"))
    except Exception as exc:
        try:
            conn.logout()
        except Exception:
            pass
        raise OutlookOtpError(f"imap XOAUTH2 login failed: {exc}") from exc
    if status != "OK":
        try:
            conn.logout()
        except Exception:
            pass
        raise OutlookOtpError(f"imap XOAUTH2 login returned {status}")
    return conn


def _decode_header(value: str) -> str:
    if not value:
        return ""
    parts: list[str] = []
    for raw, charset in email.header.decode_header(value):
        if isinstance(raw, bytes):
            parts.append(raw.decode(charset or "utf-8", errors="replace"))
        else:
            parts.append(raw)
    return "".join(parts)


def _body_from_message(msg: email.message.Message) -> str:
    chunks: list[str] = []
    for part in msg.walk():
        if part.get_content_maintype() == "multipart":
            continue
        content_type = part.get_content_type()
        if content_type not in ("text/plain", "text/html"):
            continue
        payload = part.get_payload(decode=True)
        if payload is None:
            continue
        charset = part.get_content_charset() or "utf-8"
        text = payload.decode(charset, errors="replace")
        if content_type == "text/html":
            parser = _HtmlToText()
            parser.feed(text)
            text = parser.text()
        chunks.append(text)
    return "\n".join(chunks)


def _is_hex_color_context(text: str, start: int) -> bool:
    if start > 0 and text[start - 1] == "#":
        return True
    before = text[max(0, start - 40):start]
    return bool(re.search(r"(?:color|background|bgcolor|fill|stroke)\s*[:=]\s*[\"']?#?\s*$", before, re.I))


def extract_otp(text: str, code_regex: str = r"(?<!\d)(\d{4,8})(?!\d)") -> str:
    """Extract the most likely OTP from text."""
    if not text:
        return ""
    patterns = (
        r"(?:otp|one[-\s]*time|verification|verify|code|passcode|security|login|confirm|"
        r"kode|verifikasi|验证码|驗證碼|验证|驗證)[^\d]{0,100}(\d{4,8})(?!\d)",
        r"(?<!\d)(\d{4,8})(?!\d)[^\n\r]{0,100}(?:otp|one[-\s]*time|verification|verify|code|"
        r"passcode|security|login|confirm|验证码|驗證碼|验证|驗證)",
        code_regex,
    )
    for pattern in patterns:
        try:
            matches = list(re.finditer(pattern, text, flags=re.I | re.S))
        except re.error:
            continue
        for match in reversed(matches):
            groups = match.groups() or (match.group(0),)
            for group in reversed(groups):
                code = re.sub(r"\D", "", str(group))
                if 4 <= len(code) <= 8 and not _is_hex_color_context(text, match.start()):
                    return code
    return ""


def _contains_any(value: str, needles: Sequence[str]) -> bool:
    if not needles:
        return True
    lower = value.lower()
    return any(needle.lower() in lower for needle in needles)


def message_matches(
    *,
    sender: str,
    subject: str,
    body: str,
    from_contains: Sequence[str],
    subject_contains: Sequence[str],
    body_contains: Sequence[str],
) -> bool:
    return (
        _contains_any(sender, from_contains)
        and _contains_any(subject, subject_contains)
        and _contains_any(body, body_contains)
    )


def _parse_message_timestamp(msg: email.message.Message) -> float:
    date_value = msg.get("Date") or ""
    try:
        return email.utils.parsedate_to_datetime(date_value).timestamp()
    except Exception:
        return 0.0


def _quote_mailbox(name: str) -> str:
    escaped = name.replace("\\", "\\\\").replace('"', '\\"')
    return f'"{escaped}"'


def list_matching_folders(conn: imaplib.IMAP4_SSL, requested: Sequence[str]) -> list[str]:
    """Return existing folders matching requested names, with fuzzy junk fallback."""
    status, listing = conn.list()
    if status != "OK" or not listing:
        return list(dict.fromkeys(requested))

    found: dict[str, str] = {}
    for raw in listing:
        if not raw:
            continue
        line = raw.decode("utf-8", errors="replace") if isinstance(raw, bytes) else str(raw)
        match = re.search(r'"([^"]+)"\s*$', line) or re.search(r"\s([^\s]+)\s*$", line)
        if match:
            name = match.group(1).strip('"')
            found[name.lower()] = name

    picked: list[str] = []
    for folder in requested:
        real = found.get(folder.lower())
        if real and real not in picked:
            picked.append(real)
    for key, real in found.items():
        if any(token in key for token in ("junk", "spam", "bulk")) and real not in picked:
            picked.append(real)
    if "INBOX" not in picked:
        picked.insert(0, "INBOX")
    return picked


def scan_once(
    conn: imaplib.IMAP4_SSL,
    *,
    folders: Sequence[str],
    since_ts: float,
    max_messages: int,
    code_regex: str,
    from_contains: Sequence[str],
    subject_contains: Sequence[str],
    body_contains: Sequence[str],
) -> Optional[MailHit]:
    for folder in folders:
        status, _ = conn.select(_quote_mailbox(folder), readonly=True)
        if status != "OK":
            continue
        status, data = conn.uid("SEARCH", None, "ALL")
        if status != "OK" or not data or not data[0]:
            continue
        uids = data[0].split()
        for raw_uid in reversed(uids[-max_messages:]):
            uid = raw_uid.decode("ascii", errors="replace")
            status, fetched = conn.uid("FETCH", raw_uid, "(BODY.PEEK[])")
            if status != "OK" or not fetched or not fetched[0]:
                continue
            try:
                raw_msg = fetched[0][1]
                msg = email.message_from_bytes(raw_msg)
            except Exception:
                continue
            ts = _parse_message_timestamp(msg)
            if since_ts and ts and ts < since_ts:
                continue
            sender = _decode_header(msg.get("From") or "")
            subject = _decode_header(msg.get("Subject") or "")
            body = _body_from_message(msg)
            if not message_matches(
                sender=sender,
                subject=subject,
                body=body,
                from_contains=from_contains,
                subject_contains=subject_contains,
                body_contains=body_contains,
            ):
                continue
            code = extract_otp(f"{subject}\n{body}", code_regex=code_regex)
            if code:
                return MailHit(
                    folder=folder,
                    uid=uid,
                    sender=sender,
                    subject=subject,
                    date=msg.get("Date") or "",
                    timestamp=ts,
                    code=code,
                )
    return None


def wait_for_otp(account: OutlookAccount, args: argparse.Namespace) -> MailHit:
    deadline = time.time() + max(1, args.timeout)
    since_ts = args.since_epoch if args.since_epoch else time.time() - args.newer_than
    token = ""
    token_at = 0.0
    last_error = ""

    while time.time() < deadline:
        conn: Optional[imaplib.IMAP4_SSL] = None
        try:
            if not token or time.time() - token_at > 3000:
                token = get_access_token(account, timeout=args.network_timeout)
                token_at = time.time()
            conn = imap_login(account, token, socket_timeout=args.network_timeout)
            folders = list_matching_folders(conn, args.folder)
            hit = scan_once(
                conn,
                folders=folders,
                since_ts=since_ts,
                max_messages=args.max_messages,
                code_regex=args.code_regex,
                from_contains=args.from_contains,
                subject_contains=args.subject_contains,
                body_contains=args.body_contains,
            )
            if hit:
                return hit
        except Exception as exc:
            last_error = str(exc)
        finally:
            if conn is not None:
                try:
                    conn.logout()
                except Exception:
                    pass
        time.sleep(max(1.0, args.interval))

    suffix = f" Last error: {last_error}" if last_error else ""
    raise TimeoutError(f"no matching OTP found within {args.timeout}s.{suffix}")


def _add_repeat_arg(parser: argparse.ArgumentParser, name: str, help_text: str) -> None:
    parser.add_argument(name, action="append", default=[], help=help_text)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Read an OTP from an authorized Outlook mailbox.")
    parser.add_argument("--email", help="Outlook mailbox. Defaults to OUTLOOK_EMAIL.")
    parser.add_argument("--client-id", help="OAuth client id. Defaults to OUTLOOK_CLIENT_ID.")
    parser.add_argument("--refresh-token", help="OAuth refresh token. Defaults to OUTLOOK_REFRESH_TOKEN.")
    parser.add_argument(
        "--account-file",
        help="File containing one email----password----client_id----refresh_token line. Prefer env vars for secrets.",
    )
    _add_repeat_arg(parser, "--from-contains", "Only accept messages whose From header contains this text.")
    _add_repeat_arg(parser, "--subject-contains", "Only accept messages whose subject contains this text.")
    _add_repeat_arg(parser, "--body-contains", "Only accept messages whose body contains this text.")
    parser.add_argument(
        "--allow-unfiltered",
        action="store_true",
        help="Allow scanning without sender/subject/body filters. Not recommended.",
    )
    parser.add_argument("--folder", action="append", default=list(DEFAULT_FOLDERS), help="Folder to scan.")
    parser.add_argument("--timeout", type=int, default=180, help="Seconds to wait for a matching code.")
    parser.add_argument("--interval", type=float, default=4.0, help="Polling interval in seconds.")
    parser.add_argument("--newer-than", type=int, default=600, help="Only consider messages newer than N seconds.")
    parser.add_argument("--since-epoch", type=float, default=0.0, help="Only consider messages after this epoch timestamp.")
    parser.add_argument("--max-messages", type=int, default=20, help="Recent messages to inspect per folder.")
    parser.add_argument("--network-timeout", type=int, default=20, help="HTTP/IMAP network timeout in seconds.")
    parser.add_argument("--code-regex", default=r"(?<!\d)(\d{4,8})(?!\d)", help="Regex used as final OTP fallback.")
    parser.add_argument("--json", action="store_true", help="Print a JSON object instead of only the code.")
    return parser


def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    if not args.allow_unfiltered and not (args.from_contains or args.subject_contains or args.body_contains):
        parser.error("provide at least one --from-contains/--subject-contains/--body-contains filter")
    try:
        account = account_from_env_or_args(args)
        hit = wait_for_otp(account, args)
    except Exception as exc:
        print(f"ERR: {exc}", file=sys.stderr)
        return 1

    if args.json:
        print(json.dumps(dataclasses.asdict(hit), ensure_ascii=False))
    else:
        print(hit.code)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
