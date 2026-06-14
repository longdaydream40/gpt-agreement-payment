# Gpt-Agreement-Payment 后端架构文档

> 本文档仅覆盖后端。前端 Vue 3 / WebUI 不在范围内。

---

## 1. 项目概要

一个 ChatGPT Plus/Team 订阅支付流程的端到端重放工具，逆向实现完整协议链：

```
注册 ChatGPT 账号 → Stripe Checkout → PayPal/GoPay/QRIS 支付 → Codex OAuth + PKCE → refresh_token 入库
```

**技术栈：** Python 3 + FastAPI + SQLite + Playwright/Camoufox + curl_cffi + Node.js (WhatsApp sidecar / QuickJS Sentinel)

**目录结构（仅后端）：**

```
├── webui/
│   ├── server.py                  # FastAPI 应用工厂
│   ├── backend/
│   │   ├── db.py                  # SQLite 数据库层 (9 张表)
│   │   ├── auth.py                # Cookie Session 鉴权
│   │   ├── settings.py            # 全局路径配置
│   │   ├── runner.py              # 单 pipeline 子进程控制器
│   │   ├── parallel_runner.py     # N-worker 并发控制器
│   │   ├── auto_loop.py           # 自动循环 runner
│   │   ├── wa_relay.py            # WhatsApp Web 侧车生命周期管理
│   │   ├── outlook_pool.py        # Outlook 接码池管理
│   │   ├── outlook_oauth_refresh.py # Outlook OAuth RT 续期
│   │   ├── config_writer.py       # Wizard 答案 → JSON 配置文件
│   │   ├── config_health.py       # 启动前配置健康检查
│   │   ├── account_inventory.py   # 账号库存摘要（只读）
│   │   ├── account_validator.py   # ChatGPT 账号有效性探针
│   │   ├── link_state.py          # GoPay 手机绑定状态追踪
│   │   ├── routes/                # 17 个路由模块
│   │   │   ├── setup.py           # 首次初始化
│   │   │   ├── auth.py            # 登录/登出/me
│   │   │   ├── wizard.py          # 14 步配置向导
│   │   │   ├── preflight.py       # 预检分发
│   │   │   ├── config.py          # 配置导出/健康检查
│   │   │   ├── run.py             # 单 pipeline 运行控制
│   │   │   ├── run_parallel.py    # N-worker 并发运行
│   │   │   ├── inventory.py       # 账号库存管理
│   │   │   ├── outlook.py         # Outlook 池管理
│   │   │   ├── whatsapp.py        # WhatsApp 侧车控制
│   │   │   ├── promo_links.py     # Promo 链接池
│   │   │   ├── auto_loop.py       # 自动循环控制
│   │   │   ├── proxy.py           # 代理 IP 轮换
│   │   │   ├── link_state.py      # GoPay 绑定状态
│   │   │   ├── sniff.py           # Stripe 指纹嗅探 SSE
│   │   │   ├── cloudflare_kv.py   # CF KV 管理
│   │   │   └── __init__.py
│   │   └── preflight/             # 10 个预检模块
│   └── tests/                     # pytest 测试
├── pipeline/
│   ├── __init__.py                # 从 _monolith re-export
│   ├── __main__.py                # python -m pipeline 入口
│   ├── _monolith.py               # 4022 行核心调度器
│   ├── cpa_autofill.py            # CPA 散户面板上传
│   ├── promo_link.py              # Promo 长链接抓取
│   └── oauth/
│       └── team_api.py            # Codex OAuth + Team API
├── CTF-pay/                       # 支付重放模块
│   ├── card/                      # Stripe 卡支付
│   ├── gopay/                     # GoPay 支付
│   ├── qris/                      # QRIS 扫码支付
│   ├── captcha/                   # hCaptcha 桥接 + 解算
│   ├── hcaptcha_auto_solver.py    # VLM hCaptcha 视觉解算器
│   ├── adb/                       # Android Debug Bridge 驱动
│   ├── recovery/                  # 支付拒绝重试
│   ├── relays/                    # Mock 网关 + WhatsApp 中继
│   └── scripts/                   # GoPay 一拖多脚本
├── CTF-reg/                       # 注册模块
│   ├── drivers/                   # Browser (Camoufox) + Protocol (HTTP)
│   ├── mail/                      # CF KV + Outlook 邮箱 OTP
│   ├── sentinel/                  # OpenAI Sentinel 解算器
│   └── paypal_plus/               # PayPal Plus 注册
├── core/                          # 共享工具
│   ├── jwt_decode.py              # JWT 解析
│   ├── otp_extractor.py           # OTP 提取器
│   └── otp_providers.py           # OTP 提供者工厂
├── scripts/                       # 独立脚本
└── docker/                        # Docker 入口
```

---

## 2. 入口与启动

### 2.1 WebUI (FastAPI)

```
uvicorn webui.server:create_app --factory --host 127.0.0.1 --port 8765
```

`webui/server.py:create_app()` 创建 FastAPI 实例，挂载 17 个路由模块、静态文件服务 (Vue SPA)、`/api/healthz` 端点。

### 2.2 CLI Pipeline

```
python pipeline.py --config CTF-pay/config.paypal.json --paypal
python -m pipeline ...  # 等价
```

支持模式：`single` / `batch` / `daemon` / `self_dealer` / `pay_only` / `register_only` / `free_register` / `free_backfill_rt` / `promo_link`

### 2.3 独立支付脚本

```
python -m card    # Stripe 卡支付
python -m gopay   # GoPay 支付
python -m qris    # QRIS 支付
```

---

## 3. 数据库设计 (SQLite)

路径: `output/webui.db`（可通过 `WEBUI_DATA_DIR` 环境变量覆盖），WAL 模式，外键开启。

### 3.1 表结构

#### `users` — 管理员用户
```sql
username TEXT PRIMARY KEY,
pw_hash BLOB NOT NULL,        -- bcrypt 12 rounds
created_at REAL NOT NULL
```

#### `sessions` — 登录会话
```sql
id TEXT PRIMARY KEY,           -- secrets.token_urlsafe(32)
username TEXT NOT NULL REFERENCES users(username) ON DELETE CASCADE,
created_at REAL NOT NULL,
expires_at REAL NOT NULL       -- TTL 7 天
```

#### `runtime_meta` — 通用 KV 存储
```sql
key TEXT PRIMARY KEY,
value TEXT NOT NULL,           -- 任意字符串（常为 JSON）
updated_at REAL NOT NULL
```

存储内容：`wizard_state`, `daemon_state`, `email_domain_state`, `secrets` (Cloudflare 凭证), `wa_state`, `wa_settings`, `wa_relay_token`, `wa_session_snapshot`, `gopay_link_state`

#### `registered_accounts` — 已注册 ChatGPT 账号
```sql
id INTEGER PRIMARY KEY AUTOINCREMENT,
email TEXT NOT NULL COLLATE NOCASE,
ts TEXT NOT NULL,
password TEXT DEFAULT '',
session_token TEXT DEFAULT '',
access_token TEXT DEFAULT '',
device_id TEXT DEFAULT '',
csrf_token TEXT DEFAULT '',
id_token TEXT DEFAULT '',
refresh_token TEXT DEFAULT '',
cookie_header TEXT DEFAULT '',
created_at REAL NOT NULL,
last_check_at REAL DEFAULT 0,
last_check_status TEXT DEFAULT '',     -- valid | invalid | unknown
last_check_message TEXT DEFAULT '',
last_plan_type TEXT DEFAULT ''         -- free | plus | team | pro
```
索引: `(email, id)`

#### `pipeline_results` — Pipeline 执行记录
```sql
id INTEGER PRIMARY KEY AUTOINCREMENT,
ts TEXT NOT NULL,
mode TEXT DEFAULT '',
status TEXT DEFAULT '',
error TEXT DEFAULT '',
registration_status TEXT DEFAULT '',
registration_email TEXT DEFAULT '',
registration_error TEXT DEFAULT '',
payment_status TEXT DEFAULT '',
payment_email TEXT DEFAULT '',
payment_error TEXT DEFAULT '',
domain TEXT DEFAULT '',
proxy TEXT DEFAULT '',
cpa_import TEXT DEFAULT '',
created_at REAL NOT NULL
```
索引: `(registration_email, id)`, `(payment_email, id)`

#### `card_results` — 卡支付详细记录
```sql
id INTEGER PRIMARY KEY AUTOINCREMENT,
ts TEXT NOT NULL,
status TEXT DEFAULT '',
chatgpt_email TEXT DEFAULT '',
email TEXT DEFAULT '',
session_id TEXT DEFAULT '',
channel TEXT DEFAULT '',
entity TEXT DEFAULT '',
config TEXT DEFAULT '',
error TEXT DEFAULT '',
refresh_token TEXT DEFAULT '',
team_account_id TEXT DEFAULT '',
invite_permission TEXT DEFAULT '',
team_gpt_account_pk TEXT DEFAULT '',
email_domain TEXT DEFAULT '',
created_at REAL NOT NULL
```
索引: `(chatgpt_email, session_id, id)`

#### `oauth_status` — OAuth 状态追踪
```sql
email TEXT PRIMARY KEY COLLATE NOCASE,
status TEXT NOT NULL,          -- succeeded | dead | transient_failed
ts TEXT NOT NULL,
fail_reason TEXT DEFAULT ''
```

#### `outlook_accounts` — Outlook 接码池
```sql
email TEXT PRIMARY KEY COLLATE NOCASE,
password TEXT DEFAULT '',
client_id TEXT DEFAULT '',
refresh_token TEXT NOT NULL,
status TEXT NOT NULL DEFAULT 'available',  -- available | in_use | used | dead
imported_at REAL DEFAULT 0,
claimed_at REAL DEFAULT 0,
used_at REAL DEFAULT 0,
chatgpt_email TEXT DEFAULT '',
fail_reason TEXT DEFAULT ''
```
索引: `(status, imported_at)`

#### `promo_links` — 优惠长链接池
```sql
id INTEGER PRIMARY KEY AUTOINCREMENT,
email TEXT NOT NULL COLLATE NOCASE,
checkout_url TEXT NOT NULL,             -- Stripe hosted long URL
cs_id TEXT DEFAULT '',                  -- cs_live_xxx
processor_entity TEXT DEFAULT '',       -- openai_llc / openai_ie
plan_name TEXT DEFAULT '',              -- chatgptplusplan / chatgptteamplan
promo_campaign_id TEXT DEFAULT '',      -- plus-1-month-free 等
billing_country TEXT DEFAULT '',
billing_currency TEXT DEFAULT '',
amount_due_cents INTEGER DEFAULT 0,
status TEXT NOT NULL DEFAULT 'fresh',   -- fresh | in_use | used | expired
created_at REAL NOT NULL,
used_at REAL DEFAULT 0,
raw_response TEXT DEFAULT '',
claimed_by TEXT DEFAULT '',             -- 并发 worker 原子声明
claimed_at REAL DEFAULT 0
```
索引: `(email, id)`, `(status, created_at)`

### 3.2 关键 DB 操作

| 方法 | 说明 |
|---|---|
| `get_db()` | 返回 Database 单例，读取 `WEBUI_DATA_DIR/webui.db` |
| `set_runtime_value(key, value)` | 写入 KV（UPSERT） |
| `get_runtime_json(key)` | 读取 JSON KV |
| `add_registered_account(row)` | 写入注册账号 |
| `add_pipeline_result(record)` | 写入 pipeline 执行记录 |
| `add_card_result(record)` | 写入卡支付记录 |
| `claim_next_fresh_promo_link(worker_id)` | 原子占用一条 fresh promo link（支持并发） |
| `update_account_check(id, status)` | 更新账号有效性检查结果 |
| `update_account_rt_status(id, ...)` | 更新 RT 刷新结果（含 token 轮转） |

---

## 4. 鉴权体系

### 4.1 用户鉴权

`webui/backend/auth.py` 提供 FastAPI 依赖注入：

- **`CurrentUser`** — 从 Cookie `session_id` 验证登录态，未登录抛 401
- **`current_user_optional`** — 可选鉴权，未登录返回 None

密码使用 bcrypt (12 rounds) 存储。登录失败时对不存在用户也执行 bcrypt 比较以防止时间侧信道。

### 4.2 Relay Token 鉴权

WhatsApp sidecar 和外部服务通过 `X-WA-Relay-Token` header 或 `token` query 参数鉴权。Token 由 `wa_relay.relay_token()` 生成（`secrets.token_urlsafe(32)`），存储在 `runtime_meta[wa_relay_token]`，首次运行时自动生成。

### 4.3 GoPay Link State 鉴权

- `GET` 端点：支持 session cookie 或 relay token
- `POST` 端点（写入）：仅 relay token

---

## 5. API 接口完整清单

### 5.1 初始化 (`/api/setup`)

| 方法 | 路径 | 说明 |
|---|---|---|
| `GET` | `/api/setup/status` | 检查是否已初始化（`user_count > 0`） |
| `POST` | `/api/setup` | 创建首个管理员用户 |

请求体: `{username, password}` (3-64 / 8-128 字符)

### 5.2 认证 (`/api`)

| 方法 | 路径 | 鉴权 | 说明 |
|---|---|---|---|
| `POST` | `/api/login` | 无 | 用户名+密码登录，设置 session cookie (httponly, 7天) |
| `POST` | `/api/logout` | 可选 | 删除 session，清除 cookie |
| `GET` | `/api/me` | CurrentUser | 返回当前用户 `{username}` |

### 5.3 配置向导 (`/api/wizard`)

| 方法 | 路径 | 鉴权 | 说明 |
|---|---|---|---|
| `GET` | `/api/wizard/state` | CurrentUser | 读取 14 步向导状态 |
| `POST` | `/api/wizard/state` | CurrentUser | 写入向导状态 |

向导状态存储在 `runtime_meta[wizard_state]`，结构: `{current_step: int, answers: dict}`。

### 5.4 预检 (`/api/preflight`)

| 方法 | 路径 | 鉴权 | 说明 |
|---|---|---|---|
| `POST` | `/api/preflight/{name}` | CurrentUser | 运行指定预检 |

name 取值：`system` / `cloudflare` / `cloudflare_kv` / `proxy` / `webshare` / `card` / `captcha` / `vlm` / `team_system` / `cpa`

预检返回统一格式: `{name, status: ok|warn|fail, message, missing[], blocking: bool, details, action}`

### 5.5 配置导出与健康检查 (`/api/config`)

| 方法 | 路径 | 鉴权 | 说明 |
|---|---|---|---|
| `POST` | `/api/config/export` | CurrentUser | 向导答案 → JSON 配置文件 |
| `POST` | `/api/config/health` | CurrentUser | 按运行模式做配置健康检查 |
| `GET` | `/api/config/health` | CurrentUser | 默认模式健康检查 |

**`POST /export`** 逻辑：
1. 读取 `config.*.example.json` 骨架
2. 深合并 wizard answers 覆盖
3. Plus 订阅下自动剥离 Team only 字段 (workspace_name / seat_quantity)
4. Cloudflare 凭证写入 `runtime_meta[secrets]`
5. 备份旧配置文件为 `.bak.<timestamp>`
6. 写入 `CTF-pay/config.paypal.json` 和 `CTF-reg/config.paypal-proxy.json`

**`POST /health`** 检查项：
- 支付/注册配置文件存在且可解析
- Cloudflare KV 凭证完整性（仅注册/pay-only 模式）
- 邮箱域名配置（仅注册模式）
- 支付配置完整性（PayPal 邮箱+密码 或 GoPay 国家码+手机号+PIN 或 卡号+CVC+有效期）
- WhatsApp relay 连接状态（GoPay 模式）
- pay-only 模式下账号库存可用性
- CPA 配置完整性（free_register/free_backfill_rt 模式强制）
- team_system 配置完整性（daemon 模式强制）
- free_backfill_rt 模式下 RT 待补账号检查

### 5.6 单 Pipeline 运行控制 (`/api/run`)

| 方法 | 路径 | 鉴权 | 说明 |
|---|---|---|---|
| `GET` | `/api/run/status` | CurrentUser | 当前运行状态 |
| `POST` | `/api/run/start` | CurrentUser | 启动 pipeline |
| `POST` | `/api/run/stop` | CurrentUser | 停止 pipeline |
| `POST` | `/api/run/otp` | CurrentUser | 提交 OTP |
| `GET` | `/api/run/logs?tail=500` | CurrentUser | 获取最近 N 行日志 |
| `GET` | `/api/run/stream` | CurrentUser | SSE 日志流 |
| `POST` | `/api/run/preview` | CurrentUser | 预览命令行（不执行） |
| `GET` | `/api/run/qris/state` | CurrentUser | QRIS 当前状态 |
| `GET` | `/api/run/qris/qr.png` | CurrentUser | QRIS QR 码 PNG |

**`POST /start` 请求体（StartRequest）：**

| 字段 | 类型 | 默认值 | 说明 |
|---|---|---|---|
| `mode` | string | `"single"` | 模式: `single` / `batch` / `self_dealer` / `daemon` / `free_register` / `free_backfill_rt` / `promo_link` / `no_card_plus` |
| `paypal` | bool | `true` | 走 PayPal 支付 |
| `gopay` | bool | `false` | 走 GoPay 支付 |
| `qris` | bool | `false` | 走 QRIS 支付 |
| `batch` | int | `0` | 批量次数（batch 模式） |
| `workers` | int | `3` | 并发数（batch 模式） |
| `self_dealer` | int | `0` | 自 dealer 成员数 |
| `register_only` | bool | `false` | 仅注册不支付 |
| `pay_only` | bool | `false` | 仅支付（复用已有账号） |
| `rt_only` | bool | `false` | 仅换 RT |
| `count` | int | `0` | 注册次数（free_register 模式，0=无限） |
| `promo_plan` | string | `"plus"` | promo 计划: `plus` / `team` |
| `promo_country` | string | `"ID"` | 2 位 ISO 国家码 |
| `promo_currency` | string | `"IDR"` | 3 位币种代码 |
| `promo_campaign_id` | string | `""` | 指定优惠活动 ID |
| `register_mode` | string | `"protocol"` | 注册模式: `browser` / `protocol` |
| `target_emails` | string[] | `[]` | 定向操作指定邮箱 |
| `mail_source` | string | `"outlook"` | 邮箱来源: `outlook` / `catch_all` |
| `outlook_email` | string | `""` | 指定 Outlook 邮箱 |
| `no_card_promo_link_id` | int | `0` | 指定 promo link ID（no_card_plus 模式） |
| `no_card_phone` | string | `""` | PayPal 注册手机号 |
| `no_card_sms_api_url` | string | `""` | 接码网关 URL |
| `no_card_otp_timeout` | int | `240` | OTP 超时 (秒) |
| `no_card_signup_retries` | int | `3` | PayPal 注册重试次数 |
| `no_card_node_rpa_timeout` | int | `900` | Node RPA 总超时 |
| `no_card_max_due` | int | `100` | 最大金额 (minor units) |
| `no_card_allow_already_paid` | bool | `false` | 允许已支付账号 |
| `no_card_allow_full_price` | bool | `false` | 允许全价链接 |
| `no_card_paypal_country` | string | `"US"` | PayPal 注册国家 |
| `no_card_paypal_lang` | string | `"en"` | PayPal 页面语言 |
| `no_card_inventory_mail_source` | string | `"any"` | 库存邮箱过滤: `any` / `outlook` / `catch_all` |

**运行流程：**
1. 收到 start 请求 → 先跑 `build_config_health()` 健康检查
2. 不通过则 400 返回具体原因
3. 通过后调 `runner.start()` → 拼接命令行 `xvfb-run -a python pipeline.py [args]`
4. 子进程以 `start_new_session=True` 独立 session 启动
5. 后台线程 `_drain()` 逐行读取 stdout → 环形 buffer (3000 行)
6. 自动检测日志中的 OTP 等待标记、GoPay 链接状态、QRIS artifacts

**SSE 日志流 (`GET /stream`)：**
- 先推送最近 200 行 backlog
- 每 300ms 检查新日志行推送
- OTP 挂起时周期性发送 `otp_pending` 心跳
- 进程退出时发送最后一波日志 + `done` 事件（含最终状态）

### 5.7 N-Worker 并发运行 (`/api/run/parallel`)

仅支持 `no_card_plus` 模式。

| 方法 | 路径 | 鉴权 | 说明 |
|---|---|---|---|
| `POST` | `/api/run/parallel/start` | CurrentUser | 启动 N 个 worker |
| `POST` | `/api/run/parallel/stop` | CurrentUser | SIGTERM → SIGKILL 所有 worker |
| `POST` | `/api/run/parallel/clear` | CurrentUser | 清除已退出的 worker 状态 |
| `GET` | `/api/run/parallel/status` | CurrentUser | 批量汇总状态 |
| `GET` | `/api/run/parallel/logs?worker_id=w1&since=0` | CurrentUser | 指定 worker 日志 |
| `POST` | `/api/run/parallel/phone-lock/acquire?phone=X&worker=Y` | 无 | 获取手机锁（Node RPA 自调用） |
| `POST` | `/api/run/parallel/phone-lock/release?phone=X&worker=Y` | 无 | 释放手机锁 |
| `GET` | `/api/run/parallel/phone-lock/list` | CurrentUser | 列出所有手机锁 |

**设计要点：**
- 每个 worker 独立子进程，独立 stdout drainer 线程
- 环形 buffer 各 4000 行
- phone-lock 是 OTP 临界区互斥：pre-OTP 并行，OTP 阶段串行化
- 锁 TTL 180s 防止 worker 崩溃 leak
- promo_link 通过 `claimed_by` 字段原子声明，多 worker 不抢同行
- 启动间隔 stagger_s 秒错开

### 5.8 账号库存 (`/api/inventory`)

| 方法 | 路径 | 鉴权 | 说明 |
|---|---|---|---|
| `GET` | `/api/inventory/accounts` | CurrentUser | 列出所有账号 + 摘要 |
| `POST` | `/api/inventory/accounts/check` | CurrentUser | 探活验证 + 查实时 plan |
| `POST` | `/api/inventory/accounts/refresh-rt-status` | CurrentUser | RT 刷新 + 更新状态 |
| `POST` | `/api/inventory/accounts/delete` | CurrentUser | 硬删除账号 |
| `POST` | `/api/inventory/accounts/cpa-push` | CurrentUser | 推送到 CPA 管理池 |
| `POST` | `/api/inventory/accounts/cpa-autofill-push` | CurrentUser | 推送到 CPA 散户面板 |

**`GET /accounts` 返回每个账号：**
```json
{
  "id": 1,
  "email": "user@example.com",
  "plan_tag": "plus",          // free | plus | team | pro
  "cpa_status": "ok",          // CPA 推送状态
  "cpa_pushed": true,
  "registered_at": "2025-...",
  "attempts": 3,
  "has_session_token": true,
  "has_access_token": true,
  "has_device_id": true,
  "has_refresh_token": true,
  "pay_state": "consumed",     // reusable | consumed | no_auth
  "pay_only_eligible": false,
  "rt_state": "has_rt",        // has_rt | oauth_succeeded | dead | cooldown | retryable | missing
  "can_backfill_rt": false,
  "oauth_status": "succeeded",
  "last_check_status": "valid",
  "last_plan_type": "plus"
}
```

**`POST /accounts/check`** 对每个账号执行三级探针（`account_validator.py`）：
1. refresh_token → `POST auth.openai.com/oauth/token`（最可靠）
2. access_token → `GET chatgpt.com/backend-api/me` Bearer
3. cookie → `GET /backend-api/me` with Cookie

所有探针经 gost 中继保证 IP 一致性。状态：`valid` / `invalid` / `unknown`。

同时调 `/backend-api/accounts/check/v4-2023-04-27` 取实时 subscription plan，写回 `last_plan_type`。

**`POST /accounts/cpa-autofill-push`** 把选中账号推到散户面板：
- 必经字段：email / refresh_token / access_token / id_token 齐全
- 单价 price 必传（防止误打默认价）
- 单批 ≤1000 行
- 服务端会自己 RT-refresh 做 anti-double-spend

### 5.9 Outlook 账号池 (`/api/outlook`)

| 方法 | 路径 | 鉴权 | 说明 |
|---|---|---|---|
| `POST` | `/api/outlook/import` | CurrentUser | 批量导入 4 段格式账号 |
| `GET` | `/api/outlook/list?limit=200&status=` | CurrentUser | 列出账号 + 统计 |
| `GET` | `/api/outlook/stats` | CurrentUser | 仅统计 |
| `DELETE` | `/api/outlook/{email}` | CurrentUser | 删除单个 |
| `POST` | `/api/outlook/revalidate-all` | CurrentUser | 并发全量验证 |
| `POST` | `/api/outlook/device-code/start` | CurrentUser | OAuth device-code flow Step1 |
| `POST` | `/api/outlook/device-code/poll` | CurrentUser | OAuth device-code flow Step2 |
| `POST` | `/api/outlook/refresh-rt` | CurrentUser | Playwright OAuth 重拿 RT |

**导入格式：** 每行 `email----password----client_id----refresh_token`

**状态机：** `available → in_use → used` (注册成功) / `dead` (RT 失效)

**验证流程：**
1. RT → access_token (Microsoft v2 OAuth endpoint, IMAP scope)
2. IMAP XOAUTH2 连接到 `outlook.office365.com:993`

**RT 续期 (`/refresh-rt`)：** 走 OAuth Code Flow + Playwright Firefox (~20-40s)

### 5.10 WhatsApp 侧车 (`/api/whatsapp`)

| 方法 | 路径 | 鉴权 | 说明 |
|---|---|---|---|
| `GET` | `/api/whatsapp/status` | CurrentUser | 侧车状态 |
| `POST` | `/api/whatsapp/start` | CurrentUser | 启动侧车 |
| `POST` | `/api/whatsapp/settings` | CurrentUser | 设置引擎偏好 |
| `POST` | `/api/whatsapp/stop` | CurrentUser | 停止侧车 |
| `POST` | `/api/whatsapp/logout` | CurrentUser | 登出并清除 session |
| `POST` | `/api/whatsapp/sidecar/state` | relay token | 侧车状态回写 |
| `GET` | `/api/whatsapp/latest-otp?since=&token=` | relay token | 读取最新 OTP |
| `POST` | `/api/whatsapp/ingest` | relay token | 手动注入 OTP |
| `GET` | `/api/whatsapp/ingest-info` | CurrentUser | 获取注入端点信息 |
| `GET` | `/api/whatsapp/latest-otp-session?since=` | CurrentUser | 带 session 读 OTP |

**WhatsApp 侧车生命周期：**
1. `start()` → spawn `node webui/whatsapp_relay/index.js`
2. 支持引擎：`baileys`（推荐）/ `wwebjs`
3. Session 快照：跨重启通过 SQLite (`runtime_meta[wa_session_snapshot]`) 持久化 tar.gz+base64
4. OTP 存储：SQLite `runtime_meta[wa_state]`，含 `latest` + `history`

### 5.11 Promo 链接池 (`/api/promo-links`)

| 方法 | 路径 | 鉴权 | 说明 |
|---|---|---|---|
| `GET` | `/api/promo-links/list?limit=200&status=` | CurrentUser | 列出链接 + 统计 |
| `GET` | `/api/promo-links/stats` | CurrentUser | 仅统计 |
| `POST` | `/api/promo-links/{id}/mark-used` | CurrentUser | 标记已用 |
| `POST` | `/api/promo-links/{id}/status` | CurrentUser | 设置状态 (fresh/used/expired) |
| `POST` | `/api/promo-links/{id}/convert` | CurrentUser | 转换区域（重新生成 checkout） |
| `POST` | `/api/promo-links/convert-bulk` | CurrentUser | 批量区域转换 |
| `DELETE` | `/api/promo-links/{id}` | CurrentUser | 删除单个 |
| `DELETE` | `/api/promo-links?status=used` | CurrentUser | 批量删除（仅 used/expired） |

**区域转换流程：**
1. 从库存取该 email 的 access_token
2. 调 `promo_link.fetch_promo_link()` 用目标国家/币种重建 Stripe hosted checkout
3. mode=clone: 新增一行；mode=replace: 覆盖原行
4. require_promo_hit=true 时检查 `amount_due ≤ max_promo_amount_minor`，不命中拒绝保存

**支持的 billing 区域：** ID/US/JP/GB/IE/FR/DE/ES/IT/NL/CA/AU/NZ/SG/HK/TW/KR/BR/MX/IN/TH/MY/PH/VN 等

### 5.12 自动循环 (`/api/auto-loop`)

| 方法 | 路径 | 鉴权 | 说明 |
|---|---|---|---|
| `POST` | `/api/auto-loop/start` | CurrentUser | 启动自动循环 |
| `POST` | `/api/auto-loop/stop` | CurrentUser | 停止自动循环 |
| `GET` | `/api/auto-loop/status` | CurrentUser | 循环状态 |

**循环逻辑：**
- 每次迭代 = 一次 register + gopay/paypal 支付
- 停止条件：`success_count >= target_success` 或 `consecutive_fail >= max_consec_fail`
- 错误分类与自动补救：

| 分类 | 补救措施 |
|---|---|
| `success` | success_count++, 重置连续失败 |
| `cf_429` / `proxy_dead` | Webshare IP 轮换 |
| `otp_timeout` / `linked_exhausted` / `wallet_insufficient` / `register_failed` | 跳过 |
| `already_paid` | 标记跳过，不计入失败 |
| `coupon_ineligible` | 从库存删除该 email |
| `unknown` | 计入失败 |

- 多 zone 域名轮换：reg_fail_streak 或 zone_ip_rotations 达阈值切下一个 zone

### 5.13 代理管理 (`/api/proxy`)

| 方法 | 路径 | 鉴权 | 说明 |
|---|---|---|---|
| `GET` | `/api/proxy/current` | CurrentUser | 获取当前 Webshare IP |
| `POST` | `/api/proxy/rotate-ip` | CurrentUser | 触发 IP 轮换 |

轮换流程：POST Webshare `/proxy/list/refresh/` → 轮询新 IP → swap gost 上游

### 5.14 GoPay 链接状态 (`/api/gopay/link-state`)

| 方法 | 路径 | 鉴权 | 说明 |
|---|---|---|---|
| `GET` | `/api/gopay/link-state` | session/token | 列出所有追踪手机 |
| `GET` | `/api/gopay/link-state/{phone}` | session/token | 查询单个手机 |
| `POST` | `/api/gopay/link-state/unlink` | token only | 标记为未链接 |
| `POST` | `/api/gopay/link-state/set` | session/token | 双向写入链接状态 |

### 5.15 Stripe 指纹嗅探 (`/api/sniff`)

| 方法 | 路径 | 鉴权 | 说明 |
|---|---|---|---|
| `GET` | `/api/sniff/stripe` | CurrentUser | SSE 流式返回 Stripe 运行时指纹 |

### 5.16 Cloudflare KV (`/api/cloudflare-kv`)

CF KV namespace 管理端点，用于 catch-all 邮箱的 Email Worker 配置。

---

## 6. Pipeline 调度器 (`pipeline/_monolith.py`)

4022 行核心文件，是整个系统的调度中枢。

### 6.1 核心函数

#### `pipeline()` — 全链路（注册+支付）
```
1. 选 proxy (webshare 或直连)
2. 选 email domain (DomainPool)
3. register() 子进程 → email + credentials
4. 若 pay_only → 复用数据库最近未支付账号
5. pay() 子进程 → Stripe checkout → 支付
6. RT 交换: Codex OAuth → refresh_token
7. 写入 webui.db registered_accounts + card_results
8. CPA 导入 (若启用)
```

#### `register()` — 注册子进程
- spawn `python CTF-reg/auth_flow.py` 或 Camoufox browser
- 环境变量 `WEBUI_REG_MODE=protocol|browser`
- `WEBUI_MAIL_SOURCE=outlook|catch_all` 决定 OTP 接收方式
- 返回 `{email, password, session_token, access_token, ...}`

#### `pay()` — 支付子进程
- spawn `python -m card --config ... --paypal|--gopay|--qris`
- `WEBUI_GOPAY_OTP_URL` 指向 WebUI 内部 OTP endpoint
- 返回 `{status, refresh_token, team_account_id, ...}`

#### `batch()` — 批量运行
- PayPal 模式：并行注册 + 串行支付
- 其他模式：串行 N 次

#### `daemon()` — 12 路自愈守护进程
状态机持续监控 gpt-team 可用账号数，低于目标时自动注册+支付。包含：
- Webshare IP 自动轮换
- CF DNS 配额清理
- tmpfs 孤儿恢复
- gost 中继保活
- DataDome slider 自动拖拽
- 多 zone 轮换
- 连续失败冷却保护

#### `free_register_loop()` / `free_backfill_rt_loop()` / `promo_link_loop()`
免费注册 / 补 RT / 抓优惠链接的无限循环。

### 6.2 Cloudflare 域名管理

- **`CloudflareDomainProvisioner`** — 单个 zone 的子域名增删
- **`MultiZoneDomainProvisioner`** — 多 zone 带 fallback
- **`DomainPool`** — 持久化域名池，支持 burn/cooldown 追踪和自动补货

### 6.3 Team System Client

`TeamSystemClient` — gpt-team 系统 REST 客户端：
- 登录认证
- 通过 SSE 批量导入 RT
- 统计可用账号
- 更新代理配置

---

## 7. 支付模块详解

### 7.1 Stripe 卡支付 (`CTF-pay/card/_monolith.py`)

核心流程：
1. **生成 checkout** — `POST chatgpt.com/backend-api/payments/checkout` → Stripe session
2. **确认** — Stripe.js `confirmPayment` 或 `confirmCardPayment`
3. **3DS2 认证** — 浏览器自动化完成 3D Secure
4. **轮询确认** — 检查 payment intent 状态

支持两种模式：
- `inline_payment_method_data` — 直接传卡信息
- `shared_payment_method` — 复用已保存卡

### 7.2 GoPay (`CTF-pay/gopay/_monolith.py`)

链路：Midtrans Snap 链接 → GoPay 钱包绑定 → OTP 验证 → Charge 结算

OTP 来源：
- WhatsApp Web relay → WebUI SQLite endpoint
- 手动补录 (WebUI OTP 模态框)

关键状态标记：
- `[gopay] midtrans linking ok reference=X` → 记录 merchant reference
- `[gopay] charge settled` → 标记手机为 linked
- `[gopay] midtrans linking 406` → 提前同步 linked 状态

### 7.3 QRIS (`CTF-pay/qris/_monolith.py`)

Midtrans QRIS QR 码生成 → 用户扫码 → 轮询结算

生成 artifacts: PNG 路径 / 远端 URL / Deeplink URL / reference / 过期时间，前端可实时轮询展示。

### 7.4 hCaptcha 解算 (`CTF-pay/hcaptcha_auto_solver.py`)

~4000 行独立解算器，覆盖 12 种 hCaptcha 挑战类型：
- 主路径：VLM (Vision Language Model) 视觉识别
- 回退：CLIP / OpenCV 启发式匹配
- 动作合成：Playwright 人类化鼠标轨迹

---

## 8. 注册模块详解

### 8.1 协议驱动 (`CTF-reg/drivers/protocol.py`)

纯 HTTP 注册，不启动浏览器：
1. 生成 persona（算法化假身份）
2. OpenAI Sentinel 解算（QuickJS / Pure Python / V1 Legacy）
3. 创建账号 → 邮箱 OTP 验证
4. 返回 credentials

### 8.2 浏览器驱动 (`CTF-reg/drivers/browser.py`)

Camoufox (反检测 Firefox) + Playwright 浏览器自动化注册。

### 8.3 邮箱 OTP

两个来源，严格二选一（`WEBUI_MAIL_SOURCE`）：

- **`outlook`** — Outlook IMAP XOAUTH2，扫描 INBOX/Junk/Spam，过滤 tm1.openai.com 影子邮件
- **`catch_all`** — Cloudflare Email Worker → KV → polling

### 8.4 OpenAI Sentinel 解算

三种解算器（难度递增）：
1. **QuickJS** (Node.js) — 最快，默认优先
2. **Pure Python** — 无外部依赖
3. **V1 Legacy** — 兼容旧版

---

## 9. 共享工具层

### 9.1 OTP 提取器 (`core/otp_extractor.py`)

从文本/HTML/JSON 中提取 4-8 位 OTP 验证码：
- 关键词优先：`otp`, `kode`, `verification`, `code`
- 通用正则回退
- JSON payload 支持 `issued_after` 时间戳过滤

### 9.2 OTP 提供者工厂 (`core/otp_providers.py`)

5 种 OTP 提供者：
- CLI stdin
- 文件轮询
- WhatsApp HTTP relay 轮询
- 子进程命令轮询
- 配置字典自动分发

### 9.3 JWT 解码 (`core/jwt_decode.py`)

Base64 解码 RS256 JWT payload，提取 email、plan type、token expiry（不验证签名）。

---

## 10. 外部服务集成

| 服务 | 用途 | 集成方式 |
|---|---|---|
| **OpenAI/ChatGPT** | 账号注册、checkout 创建、OAuth、Team API | HTTP (curl_cffi Chrome136 指纹) |
| **Stripe** | 支付 checkout 创建+确认+3DS | 浏览器 + HTTP |
| **PayPal** | 账单协议支付 | Camoufox / Chromium RPA |
| **Midtrans** | GoPay/QRIS 支付网关 | HTTP API |
| **WhatsApp** | GoPay OTP 接收 | Node.js sidecar (Baileys/wwebjs) |
| **Microsoft Graph/IMAP** | Outlook OTP 读取 + OAuth | IMAP XOAUTH2 + Graph API |
| **Cloudflare** | DNS 管理 + Email Worker + KV | CF API |
| **Webshare** | 代理 IP 轮换 | HTTP API |
| **gost** | SOCKS5 代理中继 | 本地子进程 |
| **hCaptcha** | 验证码解算 | VLM + CLIP + Playwright |
| **gpt-team system** | 账号批量导入 + 统计 | REST API + SSE |
| **CPA (CLIProxyAPI)** | 管理池账号导入 | HTTP POST |
| **CPA Autofill** | 散户面板卖号 | `POST /api/supplier/upload` |

---

## 11. 关键设计决策

### 11.1 子进程隔离
Pipeline 以独立 subprocess (`start_new_session=True`) 运行，WebUI 重启不杀 pipeline。stop 时先 SIGTERM 整组（`killpg`），超时 5s 后 SIGKILL。

### 11.2 SQLite 做运行时状态
所有可变状态（wizard、WA session、GoPay link、promo links）统一走 SQLite `runtime_meta` 表，不依赖文件系统临时文件。仅用户可编辑的 JSON 配置保留为独立文件。

### 11.3 OTP 临界区互斥
并发模式下，同一手机号在同一时间只能有一个 worker 在 OTP 阶段。其他 worker 排队等锁释放。锁 TTL 180s 防泄漏。

### 11.4 原子化 Promo Link 声明
多 worker 通过 `UPDATE ... WHERE status='fresh' RETURNING *` 原子声明 promo link（SQLite 3.35+），老版本走 `BEGIN IMMEDIATE` 事务兜底。

### 11.5 错误分类与自动补救
Auto-loop 从日志尾部扫描 10 类错误模式，根据类型自动执行 IP 轮换、账号清理、zone 切换等补救操作，无需人工干预。

### 11.6 多 Zone 域名轮换
注册连续失败或单 zone IP 轮换达阈值时自动切下一个 Cloudflare zone，类似 daemon 的 MultiZoneDomainProvisioner 策略。

---

## 12. 支付路径总结

| 路径 | 支付方式 | 地区 | OTP 需求 | 特殊要求 |
|---|---|---|---|---|
| `--paypal` | PayPal 账单协议 | EU (爱尔兰) | 无 (SMS 备选) | PayPal 账号+密码 |
| `--gopay` | GoPay 电子钱包 | 印尼 (IDR) | WhatsApp OTP | GoPay 手机号+PIN |
| `--qris` | QRIS 扫码 | 印尼 (IDR) | 无 | 任何印尼电子钱包 |
| `no_card_plus` | PayPal 访客 | 多国 | SMS OTP | 接码网关 |

所有路径最终产出：ChatGPT Codex OAuth `refresh_token`，入库 SQLite。
