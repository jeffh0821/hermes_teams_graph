"""Tests for teams_graph Graph subscription manager."""

import os
import sys
import pytest
from unittest.mock import AsyncMock, MagicMock, patch

from plugins.platforms.teams_graph.subscription_manager import SubscriptionManager


@pytest.fixture
def mock_client():
    client = MagicMock()
    client.post = AsyncMock(return_value={"id": "sub-1"})
    client.get = AsyncMock(return_value={"value": []})
    client._request = AsyncMock()
    client.list_joined_teams = AsyncMock(return_value=[])
    return client


@pytest.fixture
def sub_mgr(mock_client):
    return SubscriptionManager(
        graph_client=mock_client,
        notification_url="https://example.com/webhook",
        client_state="secret-123",
    )


class TestSubscriptionLifecycle:
    @pytest.mark.asyncio
    async def test_subscribe_creates_resource(self, sub_mgr, mock_client):
        result = await sub_mgr.subscribe("/chats/getAllMessages")
        mock_client.post.assert_called_once()
        call_arg = mock_client.post.call_args[0][1]
        assert call_arg["resource"] == "/chats/getAllMessages"
        assert call_arg["changeType"] == "created,updated"
        assert call_arg["clientState"] == "secret-123"
        assert "expirationDateTime" in call_arg
        assert result["id"] == "sub-1"

    @pytest.mark.asyncio
    async def test_renew_updates_expiration(self, sub_mgr, mock_client):
        mock_client._request = AsyncMock(return_value={"id": "sub-1", "expirationDateTime": "2026-01-01T00:00:00Z"})
        await sub_mgr.subscribe("/chats/getAllMessages")
        await sub_mgr.renew("sub-1")
        mock_client._request.assert_called_with(
            "PATCH", "/subscriptions/sub-1",
            body={"expirationDateTime": mock_client._request.call_args[1]["body"]["expirationDateTime"]},
        )

    @pytest.mark.asyncio
    async def test_renewal_404_recreates(self, sub_mgr, mock_client):
        """404 on renew should trigger re-creation."""
        mock_client._request = AsyncMock(side_effect=Exception("Not Found"))
        # Give the exception a .status attribute
        mock_client._request.side_effect = type("Err", (Exception,), {"status": 404})("Not Found")
        await sub_mgr.subscribe("/chats/getAllMessages")
        await sub_mgr._renew_all()
        # Should have re-created: subscribe() calls post(), so post should be called twice
        # (once for original subscribe, once for re-create)
        assert mock_client.post.call_count == 2

    @pytest.mark.asyncio
    async def test_unsubscribe_deletes_resource(self, sub_mgr, mock_client):
        await sub_mgr.subscribe("/chats/getAllMessages")
        await sub_mgr.unsubscribe("sub-1")
        mock_client._request.assert_called_with("DELETE", "/subscriptions/sub-1")
        assert "sub-1" not in sub_mgr._subscriptions

    @pytest.mark.asyncio
    async def test_unsubscribe_all_cleans_up(self, sub_mgr, mock_client):
        mock_client.post = AsyncMock(side_effect=[
            {"id": "sub-1"},
            {"id": "sub-2"},
        ])
        await sub_mgr.subscribe("/chats/getAllMessages")
        await sub_mgr.subscribe("/teams/t1/channels/getAllMessages")
        await sub_mgr.unsubscribe_all()
        assert len(sub_mgr._subscriptions) == 0


class TestSubscriptionChats:
    @pytest.mark.asyncio
    async def test_subscribe_to_chats_uses_correct_resource(self, sub_mgr, mock_client):
        mock_client.list_chats = AsyncMock(return_value=[{"id": "chat-1", "topic": "Test Chat"}])
        results = await sub_mgr.subscribe_to_chats()
        assert len(results) == 1
        mock_client.post.assert_called_once()
        assert mock_client.post.call_args[0][1]["resource"] == "/chats/chat-1/messages"

    @pytest.mark.asyncio
    async def test_subscribe_to_chats_skips_empty_ids(self, sub_mgr, mock_client):
        mock_client.list_chats = AsyncMock(return_value=[{"id": ""}, {"id": "chat-1"}])
        results = await sub_mgr.subscribe_to_chats()
        assert len(results) == 1
        assert mock_client.post.call_args[0][1]["resource"] == "/chats/chat-1/messages"

    @pytest.mark.asyncio
    async def test_subscribe_to_chats_no_chats(self, sub_mgr, mock_client):
        mock_client.list_chats = AsyncMock(return_value=[])
        results = await sub_mgr.subscribe_to_chats()
        assert len(results) == 0


class TestRenewalLoop:
    @pytest.mark.asyncio
    async def test_start_renewal_loop_creates_task(self, sub_mgr):
        await sub_mgr.start_renewal_loop()
        assert sub_mgr._renewal_task is not None

    @pytest.mark.asyncio
    async def test_stop_renewal_loop_cancels_task(self, sub_mgr):
        await sub_mgr.start_renewal_loop()
        await sub_mgr.stop_renewal_loop()
        assert sub_mgr._renewal_task is None

    @pytest.mark.asyncio
    async def test_double_start_is_idempotent(self, sub_mgr):
        await sub_mgr.start_renewal_loop()
        task = sub_mgr._renewal_task
        await sub_mgr.start_renewal_loop()
        assert sub_mgr._renewal_task is task
        await sub_mgr.stop_renewal_loop()
