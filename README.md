# Outlook Mail Reader — Shuffle SOAR App 

A configurable, generic Outlook / Exchange Online mailbox reader for [Shuffle SOAR](https://shuffler.io/) and similar automation workflows.

Reads messages from any Outlook mailbox via the **Microsoft Graph API** using **application (app-only) credentials** — no user sign-in required, no delegated permissions needed.

---

## Why This App Exists

Shuffle SOAR's built-in email actions are delegated-permission based and require an interactive user login. In security operations and IT automation scenarios you often need to read shared mailboxes, service accounts, or monitored inboxes **without a human in the loop**.

This app fills that gap:

- Authenticates entirely with Azure AD **client credentials** (app-only)
- Reads any mailbox your app registration has been granted access to
- Exposes configurable filters so Shuffle workflows can precisely select messages
- Returns structured JSON that downstream Shuffle steps can act on immediately

---

## Architecture Overview

```
┌───────────────────────────────────────┐
│           Shuffle SOAR Workflow       │
│                                       │
│  ┌─────────────────────────────────┐  │
│  │  Outlook Mail Reader (this app) │  │
│  │  - list_messages                │  │
│  │  - get_message                  │  │
│  │  - extract_with_regex           │  │
│  └──────────────┬──────────────────┘  │
└─────────────────│─────────────────────┘
                  │ HTTPS
                  ▼
     ┌────────────────────────┐
     │  Microsoft Graph API   │
     │  /users/{mailbox}/     │
     │    messages            │
     └────────────┬───────────┘
                  │
                  ▼
     ┌────────────────────────┐
     │  Exchange Online /     │
     │  Microsoft 365 Mailbox │
     └────────────────────────┘
```

**Authentication flow:**

```
App → POST /oauth2/v2.0/token (client_credentials)
    ← access_token
App → GET  /users/{mailbox}/messages?$filter=...
    ← JSON message list
```

---

## Actions

### `list_messages` — List Messages

Retrieves messages from a mailbox folder with optional filtering.

| Parameter | Required | Description |
|---|---|---|
| `tenant_id` | ✅ | Azure AD tenant ID (GUID or domain) |
| `client_id` | ✅ | App registration client ID |
| `client_secret` | ✅ | App registration client secret |
| `mailbox` | ✅ | Target mailbox UPN or object ID |
| `folder` | — | Folder name or well-known name (default: `inbox`) |
| `subject_filter` | — | Exact subject match |
| `sender_filter` | — | Filter by sender email address |
| `unread_only` | — | `true` to return only unread messages |
| `received_after` | — | ISO 8601 datetime (inclusive lower bound) |
| `received_before` | — | ISO 8601 datetime (exclusive upper bound) |
| `body_keyword` | — | Case-insensitive substring match on message body |
| `top` | — | Max messages to return, 1–1000 (default: 25) |

**Returns:** `{ "success": true, "count": N, "messages": [...] }`

---

### `get_message` — Get Message

Fetches a single message by its Graph message ID (e.g. from a `list_messages` result).

| Parameter | Required | Description |
|---|---|---|
| `tenant_id` | ✅ | Azure AD tenant ID |
| `client_id` | ✅ | App registration client ID |
| `client_secret` | ✅ | App registration client secret |
| `mailbox` | ✅ | Target mailbox UPN or object ID |
| `message_id` | ✅ | Graph message ID |

**Returns:** `{ "success": true, "message": { "id", "subject", "sender_name", "sender_address", "received_at", "is_read", "body_preview", "body" } }`

---

### `extract_with_regex` — Extract with Regex *(optional)*

Fetches a message and applies a user-supplied Python regex pattern to its plain-text body. Returns all matches as a list.

This action is **optional and modular** — you only use it when you need to extract structured data from an email body. The core retrieval and filtering logic in `list_messages` does not depend on it.

| Parameter | Required | Description |
|---|---|---|
| `tenant_id` | ✅ | Azure AD tenant ID |
| `client_id` | ✅ | App registration client ID |
| `client_secret` | ✅ | App registration client secret |
| `mailbox` | ✅ | Target mailbox UPN or object ID |
| `message_id` | ✅ | Graph message ID |
| `pattern` | ✅ | Python-compatible regex pattern |
| `flags` | — | Flag characters: `i` (IGNORECASE), `m` (MULTILINE), `s` (DOTALL) |

**Returns:** `{ "success": true, "matches": [...], "count": N }`

**Example patterns:**

| Goal | Pattern | Flags |
|---|---|---|
| Find ticket IDs like `INC0012345` | `INC\d{7}` | `i` |
| Find IP addresses | `\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b` | |
| Find order numbers | `Order[:\s]+(\w{6,12})` | `i` |
| Find all email addresses | `[\w.+-]+@[\w-]+\.[a-z]{2,}` | `i` |

---

## Supported Folder Names

The following well-known folder names are supported without a folder ID lookup:

`inbox` · `sentitems` · `deleteditems` · `drafts` · `junkemail` · `outbox` · `archive`

Any other value is treated as a folder **display name** and resolved automatically via the `mailFolders` API.

---

## Prerequisites

1. **Microsoft 365 / Exchange Online** tenant with at least one licensed mailbox.
2. **Azure AD app registration** with:
   - `Mail.Read` **application** permission (not delegated)
   - Admin consent granted for the permission
3. A running **Shuffle SOAR** instance (self-hosted or cloud).
4. Docker (if building the app image locally).

---

## Azure AD Setup

1. Go to [Azure Portal → App registrations](https://portal.azure.com/#view/Microsoft_AAD_IAM/ActiveDirectoryMenuBlade/~/RegisteredApps) → **New registration**.
2. Give it a name (e.g. `shuffle-outlook-reader`). No redirect URI needed.
3. Go to **Certificates & secrets** → **New client secret** → copy the value immediately.
4. Go to **API permissions** → **Add a permission** → **Microsoft Graph** → **Application permissions** → add `Mail.Read`.
5. Click **Grant admin consent**.
6. Note your **Tenant ID** and **Client ID** from the app overview page.

> **Security note:** Grant `Mail.Read` scope only. This app does not need `Mail.ReadWrite`, `Mail.Send`, or any other permission. Apply the principle of least privilege.

---

## Installation in Shuffle

### Option A — Import from URL (recommended)

1. In Shuffle, go to **Apps** → **New app** → **Import from URL**.
2. Enter the raw URL to `api.yaml` in this repository.

### Option B — Build and load locally

```bash
git clone https://github.com/uldagalihan/outlook-graph-app-only.git
cd outlook-graph-app-only/shuffle-outlook-mail-reader/1.0.2

docker build -t outlook-mail-reader:2.0.0 .
```

Then load the image into your Shuffle Docker environment and reference it in `api.yaml`.

---

## Configuration

Credentials are configured as **Shuffle app authentication parameters** (marked `configuration: true` in `api.yaml`). They are stored securely by Shuffle and never appear in plaintext in workflow definitions.

| Parameter | Where to find it |
|---|---|
| `tenant_id` | Azure Portal → App registration page → Overview |
| `client_id` | Azure Portal → App registration page → Overview |
| `client_secret` | Azure Portal → Certificates & secrets (copy at creation time) |

For local testing outside Shuffle, copy `.env.example` to `.env` and fill in your values:

```bash
cp .env.example .env
# Edit .env with real values (never commit .env to git)
```

---

## Example Scenarios

### 1. Poll an unread inbox for security alerts

```
Action:  list_messages
mailbox:     soc-alerts@contoso.com
folder:      inbox
unread_only: true
top:         50
```

### 2. Find all emails from a specific sender this week

```
Action:         list_messages
mailbox:        shared-inbox@contoso.com
sender_filter:  notifications@vendor.com
received_after: 2024-11-01T00:00:00Z
```

### 3. Extract ticket IDs from a helpdesk notification

```
Step 1 — get_message
  message_id: {{ list_messages.messages[0].id }}

Step 2 — extract_with_regex
  message_id: {{ list_messages.messages[0].id }}
  pattern:    INC\d{7}
  flags:      i
```

### 4. Check a folder other than the inbox

```
Action: list_messages
mailbox: monitoring@contoso.com
folder:  Archive
top:     100
```

---

## Limitations

- **Read-only** — this app only reads mail. It does not send, move, delete, or mark messages.
- **Page size** — the Graph API returns a maximum of 1000 messages per request. Pagination is not yet implemented; set `top` accordingly.
- **App-only auth** — the app must be granted `Mail.Read` **application** permission with admin consent. Delegated permissions are not supported.
- **Shared mailboxes** — supported if the app registration has been granted access. No additional configuration is required.
- **Attachments** — not yet supported. Only message metadata and plain-text body are returned.

---

## Security Notes

- **Never hardcode credentials** in `app.py`, `api.yaml`, or any committed file. Always pass them as Shuffle configuration parameters.
- The `.gitignore` in this repository excludes `.env` files. Verify this before pushing a fork.
- Client secrets should be rotated regularly in Azure AD. Update the Shuffle app configuration after rotation.
- Consider restricting the app registration to specific mailboxes using [application access policies](https://learn.microsoft.com/en-us/graph/auth-limit-mailbox-access) to enforce least-privilege access.

---

## Troubleshooting

| Symptom | Likely cause | Fix |
|---|---|---|
| `401 Unauthorized` | Invalid or expired credentials | Check tenant/client ID and secret in Shuffle app config |
| `403 Forbidden` | Missing or un-consented permission | Add `Mail.Read` app permission and grant admin consent |
| `404 Not Found` on folder | Non-existent folder name | Check spelling; folder names are case-insensitive but must match exactly |
| Empty `messages` list | Filters too restrictive | Try removing optional filters one by one |
| Regex returns no matches | Pattern or body encoding issue | Test pattern in isolation; enable `DOTALL` flag (`s`) for multi-line bodies |

---

## Contributing

See [CONTRIBUTING.md](./CONTRIBUTING.md).

---

## License

MIT — see [LICENSE](./LICENSE) (to be added).
