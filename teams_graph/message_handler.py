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

_RESOURCE_RE = re.compile(
    r"chats[/\(](?P<chat_id>[^/)]+)[/\)]/messages[/\(](?P<message_id>[^/)]+)[/\)]?"
)


class ChatMessageHandler:
    """Fetches chat messages from Graph and converts to Hermes events."""

    def __init__(
        self,
        graph_client: GraphClient,
        self_user_id: str = "",
        on_message: Optional[OnMessageCallback] = None,
    ):
        self._client = graph_client
        self._self_user_id = self_user_id
        self._on_message = on_message

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

        chat_message = self._parse_message(chat_id, msg_data)

        sender_id = chat_message.raw.get("from", {}).get("user", {}).get("id", "")
        if sender_id and sender_id == self._self_user_id:
            logger.debug("Skipping own message %s", message_id)
            return None

        event = self._to_message_event(chat_message)
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

    def _to_message_event(self, msg: TeamsChatMessage) -> MessageEvent:
        from gateway.session import SessionSource
        from gateway.config import Platform

        source = SessionSource(
            platform=Platform("teams_graph"),
            chat_id=msg.chat_id,
            chat_name=msg.chat_id[:20],
            chat_type="direct" if ":" in msg.chat_id else "group",
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