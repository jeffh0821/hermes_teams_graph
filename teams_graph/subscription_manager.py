"""Microsoft Graph change notification subscription manager.

Chat message subscriptions expire after 60 minutes (Microsoft limit).
We auto-renew at 55 minutes to prevent gaps. Expired / missing
subscriptions are re-created automatically during the renewal cycle
so the platform is self-healing across gateway restarts and token
rotation.
"""

import asyncio
import logging
from datetime import datetime, timedelta, timezone
from typing import Any, Optional, Callable

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
        on_renewal_tick: Optional[Callable[[], Any]] = None,
    ):
        self._client = graph_client
        self._notification_url = notification_url
        self._client_state = client_state
        self._subscriptions: dict[str, dict[str, Any]] = {}
        self._resources: dict[str, str] = {}
        self._renewal_task: Optional[asyncio.Task] = None
        self._consecutive_failures: dict[str, int] = {}
        self._on_renewal_tick = on_renewal_tick

    async def subscribe(
        self, resource: str = "/communications/callRecords/getAllMessages"
    ) -> dict[str, Any]:
        """Create a subscription."""
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
        self._resources[sub_id] = resource
        logger.info("Created subscription %s for %s (expires %s)", sub_id, resource, expiration)
        return result

    async def renew(self, subscription_id: str) -> dict[str, Any]:
        expiration = datetime.now(timezone.utc) + timedelta(minutes=SUBSCRIPTION_LIFETIME_MINUTES)
        result = await self._client._request(
            "PATCH", f"/subscriptions/{subscription_id}",
            body={"expirationDateTime": expiration.isoformat()},
        )
        self._subscriptions[subscription_id] = result
        self._consecutive_failures.pop(subscription_id, None)
        return result

    async def unsubscribe(self, subscription_id: str) -> None:
        await self._client._request("DELETE", f"/subscriptions/{subscription_id}")
        self._subscriptions.pop(subscription_id, None)
        self._resources.pop(subscription_id, None)

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
            await self._renew_all()

    async def _renew_all(self) -> None:
        """Renew every subscription; re-create dead ones in-place."""
        for sub_id in list(self._subscriptions):
            resource = self._resources.get(sub_id)
            try:
                await self.renew(sub_id)
            except Exception as e:
                status = _http_status(e)
                if status in (404, 410):
                    logger.warning(
                        "Subscription %s gone (%d), re-creating for %s",
                        sub_id, status, resource,
                    )
                    await self._recreate(sub_id, resource)
                elif status == 403:
                    logger.error(
                        "Subscription %s forbidden (403) — permissions may have changed. "
                        "Skipping renewal for this cycle.",
                        sub_id,
                    )
                    self._consecutive_failures[sub_id] = self._consecutive_failures.get(sub_id, 0) + 1
                else:
                    logger.error("Failed to renew subscription %s: %s", sub_id, e)
                    self._consecutive_failures[sub_id] = self._consecutive_failures.get(sub_id, 0) + 1

                if self._consecutive_failures.get(sub_id, 0) >= 3 and resource:
                    logger.warning(
                        "Subscription %s has %d consecutive failures — forcing re-creation",
                        sub_id, self._consecutive_failures[sub_id],
                    )
                    await self._recreate(sub_id, resource)

        # Run optional tick callback (e.g. presence keep-alive) every cycle
        if self._on_renewal_tick is not None:
            try:
                await self._on_renewal_tick()
            except Exception as e:
                logger.error("Renewal tick callback failed: %s", e)

    async def _recreate(self, old_sub_id: str, resource: Optional[str]) -> None:
        """Remove a dead subscription and create a fresh one."""
        self._subscriptions.pop(old_sub_id, None)
        self._resources.pop(old_sub_id, None)
        self._consecutive_failures.pop(old_sub_id, None)
        if not resource:
            return
        try:
            await self.subscribe(resource)
            logger.info("Re-created subscription for %s", resource)
        except Exception as e:
            logger.error("Failed to re-create subscription for %s: %s", resource, e)

    async def subscribe_to_chats(self) -> list[dict[str, Any]]:
        """Subscribe to all chat messages for the authenticated user."""
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


def _http_status(exc: Exception) -> Optional[int]:
    """Extract HTTP status code from an aiohttp ClientResponseError."""
    try:
        return exc.status
    except AttributeError:
        return None
