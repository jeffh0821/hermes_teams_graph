# Teams Graph Platform

Full-identity Microsoft Teams via Microsoft Graph API. Apollo appears as a normal M365 user (`Apollo.AI@peigenesis.com`) — no bot framework, no @mention gating.

**Status:** ✅ Production — connected, send/receive working, subscriptions auto-renewing.

## Architecture

```
User → Teams → Graph notification → msgraph_webhook :8646 → teams_graph handler
  → fetch message via Graph → Hermes agent → send reply via Graph → Apollo in Teams
```

## Files

| File | Purpose |
|------|---------|
| `plugin.yaml` | Declares platform, env vars |
| `__init__.py` | Exports `register(ctx)` |
| `auth.py` | Token discovery: env → M365 `tokens.enc` → device code (opt-in) |
| `graph_client.py` | Async Graph client (aiohttp) with retry, rate-limit, 401 refresh |
| `subscription_manager.py` | Webhook subscription lifecycle (60-min expiry, auto-renew at 55 min) |
| `models.py` | `TeamsUser`, `TeamsChatMessage` dataclasses |
| `message_handler.py` | Graph notification → MessageEvent (handles `chats({id})` format) |
| `adapter.py` | `BasePlatformAdapter` + all hooks + `register(ctx)` |

## Dependencies

### Python packages
- `aiohttp` — async HTTP
- `msal` — token refresh
- `cryptography` — M365 `tokens.enc` decryption

### Hermes platforms
- **`msgraph_webhook`** — core platform, port 8646. Must be enabled + publicly reachable.
- **M365 skill** — `~/.hermes/skills/m365/` for token discovery (`tokens.enc`)

### Network
- Public HTTPS endpoint (Tailscale Funnel: `apollo.tail171dd2.ts.net` → `localhost:8646/msgraph/webhook`)

### Core changes (minimal)
- `gateway/platforms/msgraph_webhook.py` — `_plugin_notification_handlers` list for plugin handler registration

## Configuration

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

## Azure App Registration Permissions

| Permission | Type | Purpose |
|---|---|---|
| `User.Read` | Delegated | Read own profile |
| `Chat.Read` | Delegated | Read messages |
| `Chat.ReadWrite` | Delegated | Send messages |
| `Chat.ReadBasic` | Delegated | List chats |
| `Chat.ReadWrite.All` | Delegated | Subscription creation |
| `Chat.Create` | Delegated | Create new chats |
| `offline_access` | Delegated | Refresh tokens |

## Token Discovery

1. `TEAMS_GRAPH_ACCESS_TOKEN` env var
2. `M365_ACCESS_TOKEN` + `M365_REFRESH_TOKEN` env vars
3. M365 skill `tokens.enc` (auto-refreshes via MSAL)
4. Device-code OAuth2 (**opt-in**, off by default)

## Future Enhancements

- **Absorb msgraph_webhook** — embed HTTP listener directly, eliminate two-platform requirement
- **Independent auth** — own token storage, no M365 skill dependency
- **Channel support** — team channel subscriptions, @mention handling
- **Adaptive Cards** — rich interactive responses
- **Multi-resource subscriptions** — auto-subscribe to all joined teams' channels

## Troubleshooting

**No response to messages:** Check `allow_all_users` or `allowed_users` is set.

**403 on subscriptions:** Token lacks Chat permissions — re-auth M365 skill with Chat scopes.

**Gateway crash on startup:** Fixed — device code is opt-in, won't block unattended startup.
