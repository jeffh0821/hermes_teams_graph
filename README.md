# hermes_teams_graph

A [Hermes Agent](https://github.com/NousResearch/hermes-agent) platform plugin that integrates Microsoft Teams as a native messaging channel via Microsoft Graph API.

Apollo appears as a normal M365 user — no bot framework, no @mention gating. Send and receive messages in DMs, group chats, and channels using the authenticated user identity.

**Status:** ✅ Production — connected, send/receive working, subscriptions self-healing, presence available.

## Architecture

```
User → Teams → Graph notification → msgraph_webhook :8646 → teams_graph handler
  → fetch message via Graph → Hermes agent → send reply via Graph → Apollo in Teams
```

## Quick Start

### 1. Prerequisites

- Hermes Agent installed
- M365 tenant with Teams license
- Azure App Registration with Graph API permissions (see below)
- Public HTTPS endpoint (e.g., Tailscale Funnel, ngrok, Cloudflare Tunnel)

### 2. Azure App Registration Permissions

| Permission | Type | Purpose |
|---|---|---|
| `User.Read` | Delegated | Read own profile |
| `Chat.Read` | Delegated | Read messages |
| `Chat.ReadWrite` | Delegated | Send messages |
| `Chat.ReadBasic` | Delegated | List chats |
| `Chat.ReadWrite.All` | Delegated | Subscription creation |
| `Chat.Create` | Delegated | Create new chats |
| `Presence.ReadWrite` | Delegated | Set availability status |
| `offline_access` | Delegated | Refresh tokens |

### 3. Install

```bash
# Copy plugin into Hermes
cp -r teams_graph ~/.hermes/hermes-agent/plugins/platforms/teams_graph

# Enable in config.yaml
hermes setup    # or configure manually
```

### 4. Configuration

```yaml
platforms:
  msgraph_webhook:
    enabled: true
    extra:
      host: "127.0.0.1"
      port: 8646
      client_state: "<openssl rand -hex 32>"
      accepted_resources: []

  teams_graph:
    enabled: true
    extra:
      client_id: "<azure-app-client-id>"
      tenant_id: "<azure-tenant-id>"
      notification_url: "https://your-tunnel/msgraph/webhook"
      client_state: "<same as msgraph_webhook>"
      allow_all_users: true     # or allowed_users: "user-id-1"
```

### 5. Token Discovery

The plugin resolves tokens in order:

1. `TEAMS_GRAPH_ACCESS_TOKEN` env var
2. `M365_ACCESS_TOKEN` + `M365_REFRESH_TOKEN` env vars
3. M365 skill `tokens.enc` (auto-refreshes via MSAL)
4. Device-code OAuth2 (opt-in, off by default)

## Features

### Messaging
- Send and receive messages as a full M365 user identity
- No @mention required — messages appear directly from Apollo.AI
- Teams markdown subset: bold, italic, inline code

### Subscription Lifecycle (Self-Healing)
- Subscriptions created per-chat on connect
- Auto-renewed via PATCH every 55 minutes (60-min Microsoft limit)
- **404/410** (expired/deleted): removed and re-created immediately
- **403** (permissions): skipped for this cycle, logged
- **3+ consecutive failures**: force re-created
- Gateway restart: all subscriptions re-created fresh on `connect()`

### Presence (Availability)
- Sets Apollo to "Available" on connect, refreshed every 55 minutes
- Sets "Offline" on disconnect
- Uses `POST /me/presence/setPresence` with 1-hour session expiry

## File Structure

```
teams_graph/
├── plugin.yaml              # Plugin manifest, env vars
├── __init__.py              # register(ctx) entry point
├── adapter.py               # BasePlatformAdapter + 7 hooks + presence
├── auth.py                  # Token discovery + refresh (MSAL, PBKDF2HMAC)
├── graph_client.py          # Async Graph client (aiohttp, retry, rate-limit)
├── subscription_manager.py  # Self-healing webhook lifecycle
├── models.py                # TeamsUser, TeamsChatMessage dataclasses
├── message_handler.py       # Notification → MessageEvent
└── README.md                # Detailed plugin documentation

tests/
├── test_teams_graph_adapter.py
├── test_teams_graph_auth.py
├── test_teams_graph_client.py
└── test_teams_graph_subscription.py
```

## Dependencies

- **Hermes platforms:** `msgraph_webhook` (core platform, port 8646)
- **Core changes:** `_plugin_notification_handlers` in `msgraph_webhook.py`
- **Python:** `aiohttp`, `msal`, `cryptography`
- **Network:** Public HTTPS endpoint for webhook delivery
- **M365 skill:** Token discovery and refresh (`~/.hermes/skills/m365/`)

## Tests

46 tests passing — auth (12), client (8), subscriptions (10), adapter (16). Run with:

```bash
pytest tests/test_teams_graph_*.py
```

## Roadmap

- [ ] Absorb `msgraph_webhook` — embed HTTP listener, eliminate two-platform requirement
- [ ] Independent auth — own token storage, no M365 skill dependency
- [ ] Channel support — team channel subscriptions, @mention handling
- [ ] Multi-resource subscriptions — auto-subscribe to all joined teams' channels

## License

Proprietary — PEI-Genesis / Apollo AI
