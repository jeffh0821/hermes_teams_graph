"""Teams Graph platform adapter — full-identity Teams via Microsoft Graph."""

from __future__ import annotations

import asyncio
import logging
import os
from typing import Any, Dict, Optional

from gateway.config import PlatformConfig, Platform
from gateway.platforms.base import (
    BasePlatformAdapter,
    MessageEvent,
    MessageType,
    SendResult,
)

from .auth import TokenProvider
from .graph_client import GraphClient
from .subscription_manager import SubscriptionManager
from .message_handler import ChatMessageHandler

logger = logging.getLogger(__name__)

PLATFORM_NAME = "teams_graph"
MAX_MESSAGE_LENGTH = 28000


class TeamsGraphAdapter(BasePlatformAdapter):
    """Full-identity Teams via Microsoft Graph API.

    Messages appear from the authenticated M365 user, not a bot.
    """

    def __init__(self, config: PlatformConfig):
        platform = Platform(PLATFORM_NAME)
        super().__init__(config, platform)
        extra = config.extra or {}

        self._tp = TokenProvider(
            client_id=extra.get("client_id"),
            tenant_id=extra.get("tenant_id"),
            client_secret=extra.get("client_secret"),
        )
        self._graph: Optional[GraphClient] = None
        self._self_user_id: str = ""
        self._chat_handler: Optional[ChatMessageHandler] = None
        self._sub_mgr: Optional[SubscriptionManager] = None
        self._notification_url = extra.get("notification_url", "")
        self._client_state = extra.get("client_state", "")
        self._chat_names: dict[str, str] = {}

    # ── Lifecycle ─────────────────────────────────────────────────────────

    async def connect(self) -> bool:
        try:
            token = await self._tp.get_token(allow_device_code=False)
            logger.info("Teams Graph authenticated")
        except Exception as e:
            logger.error("Teams Graph auth failed: %s", e)
            return False

        self._graph = GraphClient(self._tp)
        await self._graph.__aenter__()

        try:
            me = await self._graph.get_me()
            self._self_user_id = me.get("id", "")
            logger.info("Connected as %s (%s)", me.get("displayName"), me.get("userPrincipalName"))
        except Exception as e:
            logger.error("Graph connectivity check failed: %s (type=%s)", e, type(e).__name__)
            return False

        self._chat_handler = ChatMessageHandler(
            graph_client=self._graph,
            self_user_id=self._self_user_id,
            on_message=lambda ev: asyncio.create_task(self.handle_message(ev)),
        )

        if self._notification_url and self._client_state:
            self._sub_mgr = SubscriptionManager(
                graph_client=self._graph,
                notification_url=self._notification_url,
                client_state=self._client_state,
            )
            try:
                await self._sub_mgr.subscribe_to_chats()
                await self._sub_mgr.start_renewal_loop()
            except Exception as e:
                logger.warning(
                    "Failed to create Graph subscriptions: %s. "
                    "The platform can send messages but will not receive them. "
                    "Add Chat.Read.All permission to the Azure app registration.",
                    e,
                )

            # Wire webhook notifications from msgraph_webhook to our handler
            await self._register_webhook_consumer()

        self._mark_connected()
        return True

    async def disconnect(self) -> None:
        if self._sub_mgr:
            await self._sub_mgr.unsubscribe_all()
            await self._sub_mgr.stop_renewal_loop()
        if self._graph:
            await self._graph.__aexit__(None, None, None)
        self._mark_disconnected()

    async def _register_webhook_consumer(self) -> None:
        """Register with the msgraph_webhook to receive chat notifications."""
        try:
            from gateway.platforms.msgraph_webhook import _plugin_notification_handlers

            async def on_notification(notification: dict, event):
                await self._chat_handler.handle_notification(notification)

            _plugin_notification_handlers.append(on_notification)
            logger.info("Registered webhook consumer with msgraph_webhook")
        except Exception as e:
            logger.error("Failed to register webhook consumer: %s", e)

    # ── Send ──────────────────────────────────────────────────────────────

    async def send(
        self,
        chat_id: str,
        content: str,
        reply_to: Optional[str] = None,
        metadata: Optional[Dict[str, Any]] = None,
    ) -> SendResult:
        if self._graph is None:
            return SendResult(success=False, error="Not connected")
        if len(content) > MAX_MESSAGE_LENGTH:
            content = content[:MAX_MESSAGE_LENGTH - 100] + "\n\n[message truncated]"
        try:
            result = await self._graph.send_chat_message(chat_id, content)
            return SendResult(success=True, message_id=result.get("id"))
        except Exception as e:
            logger.error("Failed to send to %s: %s", chat_id, e)
            return SendResult(success=False, error=str(e))

    async def get_chat_info(self, chat_id: str) -> Dict[str, Any]:
        cached = self._chat_names.get(chat_id)
        if cached:
            return {"name": cached, "chat_id": chat_id, "type": "direct"}
        try:
            chat = await self._graph.get_chat(chat_id)
            name = chat.get("topic") or "Teams Chat"
            self._chat_names[chat_id] = name
            return {"name": name, "chat_id": chat_id, "type": chat.get("chatType", "direct")}
        except Exception:
            return {"name": chat_id, "chat_id": chat_id, "type": "direct"}


# ── Hooks ──────────────────────────────────────────────────────────────────


def _env_enablement() -> dict | None:
    """Seed PlatformConfig.extra from env vars before adapter construction."""
    client_id = os.getenv("TEAMS_GRAPH_CLIENT_ID", "").strip()
    tenant_id = os.getenv("TEAMS_GRAPH_TENANT_ID", "").strip()
    if not (client_id and tenant_id):
        return None

    seed: dict = {
        "client_id": client_id,
        "tenant_id": tenant_id,
    }
    for key in ("client_secret", "notification_url", "client_state"):
        val = os.getenv(f"TEAMS_GRAPH_{key.upper()}", "").strip()
        if val:
            seed[key] = val

    home = os.getenv("TEAMS_GRAPH_HOME_CHANNEL", "").strip()
    if home:
        seed["home_channel"] = {
            "chat_id": home,
            "name": os.getenv("TEAMS_GRAPH_HOME_CHANNEL_NAME", "Home"),
        }
    return seed


def _apply_yaml_config(yaml_cfg: dict, teams_graph_cfg: dict) -> dict | None:
    """Translate config.yaml teams_graph: keys into env vars.

    Env vars take precedence — every assignment is guarded by ``not os.getenv(...)``.
    Returns a dict of extras to merge into PlatformConfig.extra.
    """
    extras: dict = {}
    teams_graph_extra = teams_graph_cfg.get("extra", {}) if isinstance(teams_graph_cfg, dict) else {}
    # Source: both top-level and nested under extra
    merged_cfg = {**teams_graph_extra, **{k: v for k, v in (teams_graph_cfg or {}).items() if k != "extra"}}

    for key in ("client_id", "tenant_id", "client_secret", "notification_url", "client_state",
                 "allowed_users", "allow_all_users"):
        yaml_val = merged_cfg.get(key) or teams_graph_cfg.get(key)
        env_key = f"TEAMS_GRAPH_{key.upper()}"
        if yaml_val is not None and not os.getenv(env_key):
            os.environ[env_key] = str(yaml_val)
            extras[key] = yaml_val

    home = teams_graph_cfg.get("home_channel")
    if home and isinstance(home, dict) and not os.getenv("TEAMS_GRAPH_HOME_CHANNEL"):
        os.environ["TEAMS_GRAPH_HOME_CHANNEL"] = home.get("chat_id", "")
        if home.get("name"):
            os.environ["TEAMS_GRAPH_HOME_CHANNEL_NAME"] = home["name"]
        extras["home_channel"] = home

    return extras if extras else None


async def _standalone_send(
    pconfig,
    chat_id: str,
    message: str,
    *,
    thread_id: Optional[str] = None,
    media_files: Optional[list] = None,
    force_document: bool = False,
) -> Dict[str, Any]:
    """Send a message via Graph API without a live gateway adapter.

    Used by ``tools/send_message_tool._send_via_adapter`` when cron
    runs separately from the gateway.
    """
    extra = getattr(pconfig, "extra", {}) or {}
    client_id = os.getenv("TEAMS_GRAPH_CLIENT_ID") or extra.get("client_id", "")
    tenant_id = os.getenv("TEAMS_GRAPH_TENANT_ID") or extra.get("tenant_id", "")

    if not client_id or not tenant_id:
        return {"error": "TEAMS_GRAPH_CLIENT_ID and TEAMS_GRAPH_TENANT_ID required"}

    tp = TokenProvider(client_id=client_id, tenant_id=tenant_id)
    async with GraphClient(tp) as client:
        try:
            if len(message) > MAX_MESSAGE_LENGTH:
                message = message[:MAX_MESSAGE_LENGTH - 100] + "\n\n[truncated]"
            result = await client.send_chat_message(chat_id, message)
            return {"success": True, "message_id": result.get("id")}
        except Exception as e:
            return {"error": str(e)}


def _is_connected(config) -> bool:
    """Return True if the platform appears configured and connected."""
    import hermes_cli.gateway as gmod
    cid = (gmod.get_env_value("TEAMS_GRAPH_CLIENT_ID") or "").strip()
    tid = (gmod.get_env_value("TEAMS_GRAPH_TENANT_ID") or "").strip()
    return bool(cid and tid)


def check_requirements() -> bool:
    try:
        import aiohttp  # noqa: F401
        import msal     # noqa: F401
        return True
    except ImportError:
        return False


def register(ctx) -> None:
    """Plugin entry point — called by the Hermes plugin system."""
    ctx.register_platform(
        name=PLATFORM_NAME,
        label="Microsoft Teams (Graph)",
        adapter_factory=lambda cfg: TeamsGraphAdapter(cfg),
        check_fn=check_requirements,
        is_connected=_is_connected,
        required_env=["TEAMS_GRAPH_CLIENT_ID", "TEAMS_GRAPH_TENANT_ID"],
        install_hint="pip install aiohttp msal cryptography",
        env_enablement_fn=_env_enablement,
        apply_yaml_config_fn=_apply_yaml_config,
        cron_deliver_env_var="TEAMS_GRAPH_HOME_CHANNEL",
        standalone_sender_fn=_standalone_send,
        allowed_users_env="TEAMS_GRAPH_ALLOWED_USERS",
        allow_all_env="TEAMS_GRAPH_ALLOW_ALL_USERS",
        max_message_length=MAX_MESSAGE_LENGTH,
        emoji="🔄",
        platform_hint=(
            "You are chatting via Microsoft Teams (Graph API). "
            "Teams renders a subset of markdown — bold (**text**), "
            "italic (*text*), and inline code (`code`) work, but "
            "complex tables or raw HTML do not. Keep responses "
            "clear and professional."
        ),
    )
