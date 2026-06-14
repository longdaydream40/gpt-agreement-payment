# Outlook 接码：Graph 优先读取最近邮件

项目支持接码商常见的 Outlook 4 段格式：

```text
email----password----client_id----refresh_token
```

当前 WebUI 注册流程里，Outlook 池账号会经由
`CTF-reg/mail/provider.py` 调用 `webui.backend.outlook_pool.fetch_otp_via_imap()`。
这个函数名保留了历史叫法，但现在实际逻辑是 **Graph 优先，IMAP 兜底**。

## 当前链路

1. 先用 mailbox `refresh_token` 换 Microsoft Graph access token：

   ```text
   scope=https://graph.microsoft.com/Mail.Read offline_access
   ```

2. 通过 HTTPS 只读取每个目录的最近一封邮件：

   ```text
   /v1.0/me/mailFolders/inbox/messages?$top=1&$orderby=receivedDateTime desc
   /v1.0/me/mailFolders/junkemail/messages?$top=1&$orderby=receivedDateTime desc
   ```

3. 在 Inbox/Junk 的最新候选里，按时间选择最新的 OpenAI 发件人邮件，并从
   subject/body 中提取 4-8 位验证码。

4. 如果 Graph 无权限、请求失败、或最近邮件没有 OpenAI OTP，才回落到 IMAP
   XOAUTH2：`outlook.office365.com:993`。

这个做法和 `https://app.wyx66.com/` 这类邮箱管理器的快路径一致：后端优先走
Graph HTTPS，不依赖本机 993 端口，也不做宽范围 IMAP 历史扫描。

## 代码位置

- `webui/backend/outlook_pool.py`
  - `GRAPH_MAIL_SCOPE`
  - `get_outlook_access_token(..., scope=...)`
  - `fetch_otp_via_graph(...)`
  - `fetch_otp_via_imap(...)`：Graph-first 包装函数，失败后 IMAP fallback
  - `_message_body_text(...)` / `_extract_otp_from_html(...)`
- `CTF-reg/mail/provider.py`
  - Outlook 池账号等待验证码时调用 `outlook_pool.fetch_otp_via_imap(...)`
- `webui/backend/routes/outlook.py`
  - Outlook 池导入、列表、重验接口

## 注意点

- Microsoft access token 的资源由 `scope` 决定。IMAP scope 换出来的 token 不能
  直接拿去调 Graph；Graph scope 换出来的 token 也不能用于 IMAP XOAUTH2。
- 如果本机到 `outlook.office365.com:993` TLS 握手超时，Graph HTTPS 仍可能正常。
  当前代码会先走 Graph 来绕开这类 IMAP 网络问题。
- 当前读取策略是“每个目录只看最近一封”，适合 OpenAI 刚触发发信后的 OTP
  polling，能减少命中过期旧码的概率。
- OpenAI 发件人过滤当前接受：

  ```text
  openai.com
  auth.openai
  tm.openai
  chatgpt.com
  tm.open
  ```

  `tm1.openai` 会被跳过，因为它曾出现固定影子码。

## 快速检查

从 WebUI `/outlook` 导入账号，或在 Python 里直接导入：

```python
from webui.backend import outlook_pool

outlook_pool.import_lines(text, validate=False)
```

单号验证 Graph 是否能读邮箱，不打印任何敏感 token：

```python
from webui.backend import outlook_pool

otp = outlook_pool.fetch_otp_via_graph(email, refresh_token, client_id)
print(bool(otp), len(otp or ""))
```

如果没有抛错但返回空，说明 Graph 权限可用，只是 Inbox/Junk 最新邮件里暂时没有
OpenAI OTP。

## 测试数据清理

Outlook 池数据在 SQLite：

```text
output/webui.db
table: outlook_accounts
```

清理测试导入时，应按导入文件里的 email 精确删除。不要在已有真实账号池时直接
清空整张表。
