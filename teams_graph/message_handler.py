"""Processes Teams chat webhook notifications into Hermes MessageEvents."""

import asyncio
import logging
import re
from typing import Any, Optional, Callable

from gateway.platforms.base import MessageEvent, MessageType
from .graph_client import GraphClient
from .models import TeamsUser, TeamsChatMessage

logger = logging.getLogger(__name__)

OnMessageCallback = Callable[[MessageEvent], asyncio.Future | None]
OnApprovalCallback = Callable[[str, str], asyncio.Future | None]  # (chat_id, text)

_RESOURCE_RE = re.compile(
    r"chats[/\\(](?P<chat_id>[^/)]+)[/\\)]/messages[/\\(](?P<message_id>[^/)]+)[/\\)]?"
)

_APPROVAL_CMD_RE = re.compile(
    r"/(approve-once|approve-session|always-allow|deny)\s+(\S+)",
    re.IGNORECASE,
)


def _is_mentioned(msg_data: dict[str, Any], self_user_id: str) -> bool:
    """Return True if self_user_id appears in the message's mentions array."""
    if not self_user_id:
        return False
    for mention in msg_data.get("mentions", []) or []:
        mentioned_user = mention.get("mentioned", {}).get("user", {})
        if mentioned_user.get("id") == self_user_id:
            return True
    return False


class ChatMessageHandler:
    """Fetches chat messages from Graph and converts to Hermes events."""

    def __init__(
        self,
        graph_client: GraphClient,
        self_user_id: str = "",
        on_message: Optional[OnMessageCallback] = None,
        on_approval_command: Optional[OnApprovalCallback] = None,
    ):
        self._client = graph_client
        self._self_user_id = self_user_id
        self._on_message = on_message
        self._on_approval_command = on_approval_command
        self._chat_meta: dict[str, dict[str, str]] = {}  # chat_id → {type, topic}

    async def _get_chat_meta(self, chat_id: str) -> dict[str, str]:
        """Fetch and cache chat type and topic. Returns {'type': ..., 'topic': ...}."""
        if chat_id in self._chat_meta:
            return self._chat_meta[chat_id]
        try:
            chat = await self._client.get_chat(chat_id)
            meta = {
                "type": chat.get("chatType", "unknown"),
                "topic": chat.get("topic") or "",
            }
        except Exception:
            meta = {"type": "unknown", "topic": ""}
        self._chat_meta[chat_id] = meta
        return meta

    async def handle_notification(
        self, notification: dict[str, Any]
    ) -> Optional[MessageEvent]:
        """Process a Graph change notification for a chat message."""
        resource = notification.get("resource", "")
        match = _RESOURCE_RE.search(resource)
        if not match:
            return None

        chat_id = match.group("chat_id").strip("'()")
        message_id = match.group("message_id").strip("'()")

        try:
            msg_data = await self._client.get(f"/chats/{chat_id}/messages/{message_id}")
        except Exception as e:
            logger.error("Failed to fetch message %s: %s", message_id, e)
            return None

        # Check for text-based approval commands BEFORE self-message filtering.
        # Users reply with commands like /approve-once {key} to approve.
        body = msg_data.get("body", {})
        body_content = body.get("content", "")

        if self._on_approval_command and isinstance(body_content, str):
            if _APPROVAL_CMD_RE.search(body_content):
                await self._on_approval_command(chat_id, body_content)
                # Approval commands are control messages — don't forward to agent
                return None

        chat_message = self._parse_message(chat_id, msg_data)

        sender_id = chat_message.raw.get("from", {}).get("user", {}).get("id", "")
        if sender_id and sender_id == self._self_user_id:
            logger.debug("Skipping own message %s", message_id)
            return None

        # ── @mention filter for group/meeting chats ──────────────────────
        chat_meta = await self._get_chat_meta(chat_id)
        chat_type = chat_meta.get("type", "unknown")

        if chat_type not in ("oneOnOne",):
            # Group chat or meeting — only respond when @mentioned
            if not _is_mentioned(msg_data, self._self_user_id):
                logger.debug(
                    "Skipping unmentioned message in %s chat %s",
                    chat_type, chat_meta.get("topic", chat_id)[:40],
                )
                return None

        event = self._to_message_event(chat_message, chat_meta)
        if self._on_message:
            await self._on_message(event)
        return event

    def _parse_message(self, chat_id: str, data: dict[str, Any]) -> TeamsChatMessage:
        sender_data = data.get("from", {}).get("user", {})
        sender = TeamsUser.from_graph(sender_data) if sender_data else None
        body = data.get("body", {})
        return TeamsChatMessage(
            id=data["id"],
            chat_id=chat_id,
            content=body.get("content", ""),
            content_type=body.get("contentType", "text"),
            sender=sender,
            created_at=data.get("createdDateTime"),
            raw=data,
        )

    def _to_message_event(self, msg: TeamsChatMessage, chat_meta: dict[str, str] | None = None) -> MessageEvent:
        from gateway.session import SessionSource
        from gateway.config import Platform

        if chat_meta is None:
            chat_meta = {}

        chat_topic = chat_meta.get("topic") or ""
        chat_type = chat_meta.get("type", "unknown")

        # Map Graph chatType to SessionSource chat_type
        type_map = {"oneOnOne": "direct", "group": "group", "meeting": "group"}
        session_chat_type = type_map.get(chat_type, "group")

        # Build a human-readable chat name
        if chat_topic:
            chat_name = chat_topic
        elif chat_type == "oneOnOne":
            chat_name = f"DM: {msg.sender.display_name if msg.sender else 'User'}"
        elif chat_type == "meeting":
            chat_name = "Meeting Chat"
        else:
            chat_name = "Group Chat"

        source = SessionSource(
            platform=Platform("teams_graph"),
            chat_id=msg.chat_id,
            chat_name=chat_name,
            chat_type=session_chat_type,
            user_id=msg.sender.id if msg.sender else "unknown",
            user_name=msg.sender.display_name if msg.sender else "Unknown",
        )

        return MessageEvent(
            text=msg.content,
            message_type=MessageType.TEXT,
            source=source,
            raw_message=msg.raw,
            message_id=msg.id,
        )