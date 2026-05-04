"""Tests for the /api/proxy routes (Webshare IP rotation control)."""
from __future__ import annotations

import json


def _login(client):
    client.post("/api/setup", json={"username": "admin", "password": "hunter2hunter2"})
    client.post("/api/login", json={"username": "admin", "password": "hunter2hunter2"})


def _write_pay_cfg(monkeypatch, tmp_path, ws_block: dict | None):
    cfg = {}
    if ws_block is not None:
        cfg["webshare"] = ws_block
    p = tmp_path / "config.paypal.json"
    p.write_text(json.dumps(cfg), encoding="utf-8")
    import webui.backend.settings as settings_mod
    monkeypatch.setattr(settings_mod, "PAY_CONFIG_PATH", p)


def test_proxy_current_requires_auth(client):
    r = client.get("/api/proxy/current")
    assert r.status_code == 401


def test_proxy_current_400_when_disabled(client, tmp_path, monkeypatch):
    _login(client)
    _write_pay_cfg(monkeypatch, tmp_path, {"enabled": False})
    r = client.get("/api/proxy/current")
    assert r.status_code == 400
    assert "未启用" in r.json()["detail"]


def test_proxy_current_400_when_no_api_key(client, tmp_path, monkeypatch):
    _login(client)
    _write_pay_cfg(monkeypatch, tmp_path, {"enabled": True, "api_key": ""})
    r = client.get("/api/proxy/current")
    assert r.status_code == 400
    assert "api_key" in r.json()["detail"]


def test_proxy_rotate_requires_auth(client):
    r = client.post("/api/proxy/rotate-ip")
    assert r.status_code == 401


def test_proxy_rotate_400_when_disabled(client, tmp_path, monkeypatch):
    _login(client)
    _write_pay_cfg(monkeypatch, tmp_path, {"enabled": False})
    r = client.post("/api/proxy/rotate-ip")
    assert r.status_code == 400


def test_proxy_rotate_calls_pipeline_helper(client, tmp_path, monkeypatch):
    """Smoke: with valid config, /rotate-ip ends up calling _rotate_webshare_ip
    in pipeline.py. We monkeypatch the helper to avoid real HTTP and verify
    the response shape."""
    _login(client)
    _write_pay_cfg(monkeypatch, tmp_path, {"enabled": True, "api_key": "fake"})

    import sys
    sys.path.insert(0, "/home/kali/gpt-mitm-public")
    import pipeline as pl

    fake_px = {
        "proxy_address": "1.2.3.4",
        "port": 9999,
        "country_code": "ID",
        "asn_name": "Fake ASN",
        "valid": True,
    }
    monkeypatch.setattr(pl, "_rotate_webshare_ip", lambda *a, **kw: fake_px)
    monkeypatch.setattr(
        pl.WebshareClient, "get_current_proxy",
        lambda self: {"proxy_address": "9.9.9.9"},
    )

    r = client.post("/api/proxy/rotate-ip")
    assert r.status_code == 200
    body = r.json()
    assert body["ok"] is True
    assert body["prev_ip"] == "9.9.9.9"
    assert body["new_ip"] == "1.2.3.4"
    assert body["country"] == "ID"


def test_proxy_rotate_returns_402_on_quota_exhausted(client, tmp_path, monkeypatch):
    _login(client)
    _write_pay_cfg(monkeypatch, tmp_path, {"enabled": True, "api_key": "fake"})

    import sys
    sys.path.insert(0, "/home/kali/gpt-mitm-public")
    import pipeline as pl

    def boom(*a, **kw):
        raise pl.WebshareQuotaExhausted("quota done")

    monkeypatch.setattr(pl, "_rotate_webshare_ip", boom)
    monkeypatch.setattr(
        pl.WebshareClient, "get_current_proxy",
        lambda self: {"proxy_address": "9.9.9.9"},
    )

    r = client.post("/api/proxy/rotate-ip")
    assert r.status_code == 402
    assert "quota" in r.json()["detail"].lower() or "额度" in r.json()["detail"]
