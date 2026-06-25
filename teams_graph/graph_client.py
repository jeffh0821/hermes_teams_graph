"""Async Microsoft Graph API client for Teams chat operations."""

from __future__ import annotations

import asyncio
import logging
from typing import Any, Optional
from urllib.parse import urljoin

try:
    import aiohttp
    AIOHTTP_AVAILABLE = True
except ImportError:
    AIOHTTP_AVAILABLE = False

from .auth import TokenProvider

logger = logging.getLogger(__name__)

GRAPH_BASE = "https://graph.microsoft.com/v1.0/"
DEFAULT_RETRIES = 3
BASE_DELAY = 1.0


class GraphClient:
    """Async client for Microsoft Graph API with retry + rate-limit handling."""

    def __init__(self, token_provider: TokenProvider):
        self._tp = token_provider
        self._session: Optional[aiohttp.ClientSession] = None

    async def __aenter__(self):
        self._session = aiohttp.ClientSession()
        return self

    async def __aexit__(self, *args):
        if self._session:
            await self._session.close()
            self._session = None

    async def _headers(self) -> dict[str, str]:
        token = await self._tp.get_token()
        return {
            "Authorization": f"Bearer {token.access_token}",
            "Content-Type": "application/json",
        }

    async def _request(
        self,
        method: str,
        path: str,
        body: Optional[dict] = None,
        params: Optional[dict] = None,
    ) -> dict[str, Any]:
        url = urljoin(GRAPH_BASE, path.lstrip("/"))
        headers = await self._headers()
        last_exc = None

        for attempt in range(DEFAULT_RETRIES):
            try:
                async with self._session.request(
                    method, url, headers=headers, json=body, params=params
                ) as resp:
                    if resp.status == 429:
                        retry_after = int(resp.headers.get("Retry-After", BASE_DELAY * (2 ** attempt)))
                        logger.warning("Graph rate limited, waiting %ds", retry_after)
                        await asyncio.sleep(retry_after)
                        continue
                    if resp.status == 401:
                        if await self._tp._refresh_if_needed():
                            headers = await self._headers()
                            continue
                    if resp.status >= 500 and attempt < DEFAULT_RETRIES - 1:
                        delay = BASE_DELAY * (2 ** attempt)
                        await asyncio.sleep(delay)
                        continue
                    resp.raise_for_status()
                    # Read body INSIDE the async with — the connection is still open
                    # 204 No Content has no body (common for DELETE operations)
                    if resp.status == 204:
                        return {}
                    return await resp.json()
            except aiohttp.ClientResponseError as e:
                last_exc = e
                if e.status >= 500 and attempt < DEFAULT_RETRIES - 1:
                    delay = BASE_DELAY * (2 ** attempt)
                    await asyncio.sleep(delay)
                    continue
                raise
        raise last_exc or RuntimeError(f"Graph request failed after {DEFAULT_RETRIES} retries")

    async def get(self, path: str, **params) -> dict[str, Any]:
        return await self._request("GET", path, params=params)

    async def post(self, path: str, body: dict[str, Any]) -> dict[str, Any]:
        return await self._request("POST", path, body=body)

    # ── Convenience methods ─────────────────────────────────────────────────

    async def get_me(self) -> dict[str, Any]:
        return await self.get("/me")

    async def list_chats(self) -> list[dict[str, Any]]:
        return (await self.get("/chats")).get("value", [])

    async def get_chat(self, chat_id: str) -> dict[str, Any]:
        return await self.get(f"/chats/{chat_id}")

    async def get_chat_messages(self, chat_id: str, top: int = 10) -> list[dict[str, Any]]:
        return (await self.get(f"/chats/{chat_id}/messages", **{"$top": top})).get("value", [])

    async def send_chat_message(
        self, chat_id: str, content: str, content_type: str = "text"
    ) -> dict[str, Any]:
        return await self.post(f"/chats/{chat_id}/messages", {
            "body": {"content": content, "contentType": content_type}
        })

    async def list_joined_teams(self) -> list[dict[str, Any]]:
        return (await self.get("/me/joinedTeams")).get("value", [])

    async def get_team_channels(self, team_id: str) -> list[dict[str, Any]]:
        return (await self.get(f"/teams/{team_id}/channels")).get("value", [])

    async def send_channel_message(
        self, team_id: str, channel_id: str, content: str
    ) -> dict[str, Any]:
        return await self.post(f"/teams/{team_id}/channels/{channel_id}/messages", {
            "body": {"content": content, "contentType": "text"}
        })

    async def send_chat_card(
        self, chat_id: str, card_json: dict[str, Any], fallback_text: str = ""
    ) -> dict[str, Any]:
        """Send a message containing an Adaptive Card attachment.

        The card is delivered as an attachment on the chat message.
        ``fallback_text`` is used as the HTML body content for clients
        that don't render Adaptive Cards.
        """
        import json as _json

        card_str = _json.dumps(card_json)
        card_id = f"card_{abs(hash(card_str)) % 100000}"

        html_fallback = (
            f"<attachment id=\"{card_id}\"></attachment>"
            if not fallback_text
            else fallback_text
        )

        return await self.post(f"/chats/{chat_id}/messages", {
            "body": {
                "contentType": "html",
                "content": html_fallback,
            },
            "attachments": [{
                "id": card_id,
                "contentType": "application/vnd.microsoft.card.adaptive",
                "content": card_str,
            }],
        })