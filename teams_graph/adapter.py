"""Teams Graph platform adapter — full-identity Teams via Microsoft Graph."""

from __future__ import annotations

import asyncio
import logging
import os
from typing import Any, Dict, Optional

import asyncio as _asyncio
import json as _json
import os as _os
import re as _re

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
            on_approval_command=lambda chat_id, text: asyncio.create_task(
                self.handle_approval_command(chat_id, text)
            ),
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

    @staticmethod
    def _format_markdown_to_html(text: str) -> str:
        """Convert basic markdown to HTML suitable for Teams Graph API.

        Teams renders HTML in chat messages sent via Graph API with
        ``contentType: html``.  This converter handles the subset of
        markdown that the LLM is instructed to produce (bold, italic,
        code, line breaks, paragraphs, bullet lists).
        """
        import html as _html

        out = text

        # Code blocks first (preserve content inside them)
        code_blocks = []
        def _save_cb(m):
            code_blocks.append(m.group(1))
            return f"\x00CODEBLOCK{len(code_blocks)-1}\x00"
        out = _re.sub(r"```(.*?)```", _save_cb, out, flags=_re.DOTALL)

        # Inline code
        inline_codes = []
        def _save_ic(m):
            inline_codes.append(m.group(1))
            return f"\x00INLINECODE{len(inline_codes)-1}\x00"
        out = _re.sub(r"`([^`]+)`", _save_ic, out)

        # Bold
        out = _re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", out)

        # Italic
        out = _re.sub(r"\*(.+?)\*", r"<i>\1</i>", out)

        # HTML-escape remaining text (but not our markers or HTML tags)
        parts = _re.split(r"(\x00CODEBLOCK\d+\x00|\x00INLINECODE\d+\x00|<[^>]+>)", out)
        for i, part in enumerate(parts):
            if not part:
                continue
            if part.startswith("\x00"):
                continue
            if part.startswith("<") and part.endswith(">"):
                continue
            parts[i] = _html.escape(part)
        out = "".join(parts)

        # Restore code blocks
        for i, cb in enumerate(code_blocks):
            escaped = _html.escape(cb)
            out = out.replace(f"\x00CODEBLOCK{i}\x00", f"<pre>{escaped}</pre>")

        # Restore inline code
        for i, ic in enumerate(inline_codes):
            escaped = _html.escape(ic)
            out = out.replace(f"\x00INLINECODE{i}\x00", f"<code>{escaped}</code>")

        # Paragraphs: double newlines
        out = _re.sub(r"\n\n+", "</p><p>", out)

        # Single newlines → <br>
        out = out.replace("\n", "<br>")

        # Wrap in paragraph tags
        out = f"<p>{out}</p>"

        # Clean up empty paragraphs
        out = out.replace("<p></p>", "")
        out = _re.sub(r"<p>(\s*<br>\s*)+</p>", "", out)

        return out

    async def send(
        self,
        chat_id: str,
        content: str,
        reply_to: Optional[str] = None,
        metadata: Optional[Dict[str, Any]] = None,
    ) -> SendResult:
        if self._graph is None:
            return SendResult(success=False, error="Not connected")

        # Convert markdown to HTML for Teams rendering
        html = self._format_markdown_to_html(content)
        html_len = len(html)

        if html_len > MAX_MESSAGE_LENGTH:
            html = html[:MAX_MESSAGE_LENGTH - 100] + "<br><br>[message truncated]"

        try:
            result = await self._graph.send_chat_message(chat_id, html, content_type="html")
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

    # ── Adaptive Card Approvals ──────────────────────────────────────────

    # Reply command patterns for text-based approval
    _APPROVAL_COMMANDS = [
        ("/approve-once", "approve_once", "once"),
        ("/approve-session", "approve_session", "session"),
        ("/always-allow", "approve_always", "always"),
        ("/deny", "deny", "deny"),
    ]

    _APPROVAL_CMD_RE = _re.compile(
        r"/(approve-once|approve-session|always-allow|deny)\s+(\S+)",
        _re.IGNORECASE,
    )

    async def send_exec_approval(
        self,
        chat_id: str,
        command: str,
        session_key: str,
        description: str = "dangerous command",
        metadata: Optional[Dict[str, Any]] = None,
    ) -> SendResult:
        """Send an Adaptive Card approval prompt.

        Teams Graph API cards don't support ``Action.Submit`` or
        ``Action.Execute`` without a Bot Framework registration.
        Instead, the card displays the approval request and uses
        ``Action.ShowCard`` to reveal text-reply instructions.

        The user replies with one of these commands:
            /approve-once {key}
            /approve-session {key}
            /always-allow {key}
            /deny {key}

        The message handler detects these commands and routes them
        to ``handle_approval_command()``.
        """
        if self._graph is None:
            return SendResult(success=False, error="Not connected")

        cmd_preview = command[:2000] + "..." if len(command) > 2000 else command

        card = {
            "type": "AdaptiveCard",
            "version": "1.4",
            "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
            "body": [
                {
                    "type": "TextBlock",
                    "text": "⚠️ Command Approval Required",
                    "wrap": True,
                    "weight": "Bolder",
                    "size": "Medium",
                },
                {
                    "type": "TextBlock",
                    "text": f"```\n{cmd_preview}\n```",
                    "wrap": True,
                    "fontType": "Monospace",
                    "spacing": "Medium",
                },
                {
                    "type": "TextBlock",
                    "text": f"Reason: {description}",
                    "wrap": True,
                    "isSubtle": True,
                    "spacing": "Small",
                },
            ],
            "actions": [
                {
                    "type": "Action.ShowCard",
                    "title": "📋 How to Respond",
                    "card": {
                        "type": "AdaptiveCard",
                        "body": [
                            {
                                "type": "TextBlock",
                                "text": "Reply with ONE of these commands:",
                                "wrap": True,
                                "weight": "Bolder",
                            },
                            {
                                "type": "TextBlock",
                                "text": f"/approve-once {session_key}",
                                "wrap": True,
                                "spacing": "Small",
                                "fontType": "Monospace",
                            },
                            {
                                "type": "TextBlock",
                                "text": f"/approve-session {session_key}",
                                "wrap": True,
                                "spacing": "Small",
                                "fontType": "Monospace",
                            },
                            {
                                "type": "TextBlock",
                                "text": f"/always-allow {session_key}",
                                "wrap": True,
                                "spacing": "Small",
                                "fontType": "Monospace",
                            },
                            {
                                "type": "TextBlock",
                                "text": f"/deny {session_key}",
                                "wrap": True,
                                "spacing": "Small",
                                "fontType": "Monospace",
                            },
                        ],
                    },
                },
            ],
        }

        try:
            result = await self._graph.send_chat_card(chat_id, card)
            return SendResult(success=True, message_id=result.get("id"))
        except Exception as e:
            logger.error("send_exec_approval failed: %s", e, exc_info=True)
            return SendResult(success=False, error=str(e), retryable=True)

    async def handle_approval_command(
        self, chat_id: str, text: str
    ) -> Optional[SendResult]:
        """Process a text-based approval reply command.

        Called by the message handler when an incoming message matches
        an approval command pattern (e.g. ``/approve-once abc123``).
        """
        match = self._APPROVAL_CMD_RE.search(text)
        if not match:
            return None

        cmd_name = match.group(1).lower()
        session_key = match.group(2)

        choice = None
        hermes_action = None
        for prefix, action, ch in self._APPROVAL_COMMANDS:
            if cmd_name == prefix.lstrip("/"):
                hermes_action = action
                choice = ch
                break

        if not choice or not hermes_action:
            return None

        from tools.approval import has_blocking_approval, resolve_gateway_approval

        if not has_blocking_approval(session_key):
            if self._graph:
                await self._graph.send_chat_message(
                    chat_id, "⚠️ That approval has already been resolved or expired."
                )
            return SendResult(success=True, message_id="expired")

        resolve_gateway_approval(session_key, choice)

        label_map = {
            "once": "✅ Allowed (once)",
            "session": "✅ Allowed (session)",
            "always": "✅ Always allowed",
            "deny": "❌ Denied",
        }
        result_text = label_map[choice]

        if self._graph:
            # Send a simple confirmation — no need for a card here
            await self._graph.send_chat_message(chat_id, result_text)

        return SendResult(success=True, message_id=session_key)


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
