"""Microsoft Graph change notification subscription manager.

Chat message subscriptions expire after 60 minutes (Microsoft limit).
We auto-renew at 55 minutes to prevent gaps.
"""

import asyncio
import logging
from datetime import datetime, timedelta, timezone
from typing import Any, Optional

from .graph_client import GraphClient

logger = logging.getLogger(__name__)

SUBSCRIPTION_LIFETIME_MINUTES = 60
RENEWAL_BUFFER_MINUTES = 5


class SubscriptionManager:
    """Creates and maintains Graph subscriptions for Teams chat messages."""

    def __init__(
        self,
        graph_client: GraphClient,
        notification_url: str,
        client_state: str,
    ):
        self._client = graph_client
        self._notification_url = notification_url
        self._client_state = client_state
        self._subscriptions: dict[str, dict[str, Any]] = {}
        self._renewal_task: Optional[asyncio.Task] = None

    async def subscribe(
        self, resource: str = "/communications/callRecords/getAllMessages"
    ) -> dict[str, Any]:
        """Create a subscription.

        Common resources:
        - ``/chats/getAllMessages`` — all chats the user is in
        - ``/teams/{id}/channels/getAllMessages`` — all channels in a team
        """
        expiration = datetime.now(timezone.utc) + timedelta(minutes=SUBSCRIPTION_LIFETIME_MINUTES)

        body = {
            "changeType": "created,updated",
            "notificationUrl": self._notification_url,
            "resource": resource,
            "expirationDateTime": expiration.isoformat(),
            "clientState": self._client_state,
            "latestSupportedTlsVersion": "v1_3",
        }
        result = await self._client.post("/subscriptions", body)
        sub_id = result["id"]
        self._subscriptions[sub_id] = result
        logger.info("Created subscription %s for %s (expires %s)", sub_id, resource, expiration)
        return result

    async def renew(self, subscription_id: str) -> dict[str, Any]:
        expiration = datetime.now(timezone.utc) + timedelta(minutes=SUBSCRIPTION_LIFETIME_MINUTES)
        result = await self._client.post(
            f"/subscriptions/{subscription_id}",
            {"expirationDateTime": expiration.isoformat()},
        )
        self._subscriptions[subscription_id] = result
        return result

    async def unsubscribe(self, subscription_id: str) -> None:
        await self._client._request("DELETE", f"/subscriptions/{subscription_id}")
        self._subscriptions.pop(subscription_id, None)

    async def unsubscribe_all(self) -> None:
        for sub_id in list(self._subscriptions):
            await self.unsubscribe(sub_id)

    async def start_renewal_loop(self) -> None:
        if self._renewal_task is not None:
            return
        self._renewal_task = asyncio.create_task(self._renewal_loop())

    async def stop_renewal_loop(self) -> None:
        if self._renewal_task:
            self._renewal_task.cancel()
            self._renewal_task = None

    async def _renewal_loop(self) -> None:
        interval = (SUBSCRIPTION_LIFETIME_MINUTES - RENEWAL_BUFFER_MINUTES) * 60
        while True:
            await asyncio.sleep(interval)
            for sub_id in list(self._subscriptions):
                try:
                    await self.renew(sub_id)
                except Exception as e:
                    logger.error("Failed to renew subscription %s: %s", sub_id, e)

    async def subscribe_to_chats(self) -> list[dict[str, Any]]:
        """Subscribe to all chat messages for the authenticated user.

        Per-chat subscriptions are required for delegated user context —
        the tenant-wide ``/chats/getAllMessages`` resource requires
        application permissions.
        """
        results = []
        chats = await self._client.list_chats()
        for chat in chats:
            chat_id = chat.get("id", "")
            if not chat_id:
                continue
            try:
                result = await self.subscribe(f"/chats/{chat_id}/messages")
                results.append(result)
                logger.info("Subscribed to chat %s", chat.get("topic", chat_id))
            except Exception as e:
                logger.error("Failed to subscribe to chat %s: %s",
                             chat.get("topic", chat_id), e)
        return results

    async def subscribe_to_all_joined_teams(self) -> list[dict[str, Any]]:
        """Subscribe to channel messages for all teams the user has joined."""
        results = []
        teams = await self._client.list_joined_teams()
        for team in teams:
            try:
                result = await self.subscribe(
                    f"/teams/{team['id']}/channels/getAllMessages"
                )
                results.append(result)
            except Exception as e:
                logger.error("Failed to subscribe to team %s: %s", team.get("displayName"), e)
        return results
