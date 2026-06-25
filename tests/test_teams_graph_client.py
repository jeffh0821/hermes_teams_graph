"""Tests for teams_graph Graph API client."""

import os
import sys
import pytest
from unittest.mock import AsyncMock, MagicMock, patch

from plugins.platforms.teams_graph.auth import GraphToken
from plugins.platforms.teams_graph.graph_client import GraphClient


@pytest.fixture
def token_provider():
    tp = MagicMock()
    tp.get_token = AsyncMock(return_value=GraphToken(access_token="test-token"))
    tp._refresh_if_needed = AsyncMock(return_value=True)
    return tp


class TestGraphClientHeaders:
    @pytest.mark.asyncio
    async def test_headers_include_bearer_token(self, token_provider):
        client = GraphClient(token_provider)
        client._session = MagicMock()
        headers = await client._headers()
        assert headers["Authorization"] == "Bearer test-token"
        assert headers["Content-Type"] == "application/json"


class TestGraphClientConvenience:
    @pytest.mark.asyncio
    async def test_get_me_url(self, token_provider):
        client = GraphClient(token_provider)
        client._session = MagicMock()
        with patch.object(client, 'get', new_callable=AsyncMock, return_value={"id": "user1", "displayName": "Test"}) as mock_get:
            result = await client.get_me()
            mock_get.assert_called_once_with("/me")
            assert result["id"] == "user1"

    @pytest.mark.asyncio
    async def test_list_chats_url(self, token_provider):
        client = GraphClient(token_provider)
        client._session = MagicMock()
        with patch.object(client, 'get', new_callable=AsyncMock, return_value={"value": []}) as mock_get:
            result = await client.list_chats()
            mock_get.assert_called_once_with("/chats")
            assert result == []

    @pytest.mark.asyncio
    async def test_get_chat_url(self, token_provider):
        client = GraphClient(token_provider)
        client._session = MagicMock()
        with patch.object(client, 'get', new_callable=AsyncMock, return_value={"id": "c1", "topic": "Test"}) as mock_get:
            result = await client.get_chat("c1")
            mock_get.assert_called_once_with("/chats/c1")
            assert result["topic"] == "Test"

    @pytest.mark.asyncio
    async def test_send_chat_message_url(self, token_provider):
        client = GraphClient(token_provider)
        client._session = MagicMock()
        with patch.object(client, 'post', new_callable=AsyncMock, return_value={"id": "msg1"}) as mock_post:
            result = await client.send_chat_message("c1", "hello")
            mock_post.assert_called_once()
            call_args = mock_post.call_args[0]
            assert call_args[0] == "/chats/c1/messages"
            assert call_args[1]["body"]["content"] == "hello"

    @pytest.mark.asyncio
    async def test_list_joined_teams_url(self, token_provider):
        client = GraphClient(token_provider)
        client._session = MagicMock()
        with patch.object(client, 'get', new_callable=AsyncMock, return_value={"value": []}) as mock_get:
            await client.list_joined_teams()
            mock_get.assert_called_once_with("/me/joinedTeams")


class TestGraphClientRateLimit:
    @pytest.mark.asyncio
    async def test_429_triggers_retry(self, token_provider):
        client = GraphClient(token_provider)

        class FakeResponse:
            status = 200
            headers = {"Retry-After": "0"}
            async def __aenter__(self):
                return self
            async def __aexit__(self, *args):
                pass
            async def json(self):
                return {"value": "ok"}
            def raise_for_status(self):
                pass

        class Fake429:
            status = 429
            headers = {"Retry-After": "0"}
            async def __aenter__(self):
                return self
            async def __aexit__(self, *args):
                pass
            async def json(self):
                return {}

        call_count = [0]

        class FakeSession:
            async def __aenter__(self):
                return self
            async def __aexit__(self, *args):
                pass

            def request(self, *args, **kwargs):
                call_count[0] += 1
                if call_count[0] == 1:
                    return Fake429()
                return FakeResponse()

        client._session = FakeSession()
        result = await client.get("/test")
        assert result == {"value": "ok"}
        assert call_count[0] == 2


class TestGraphClientRetry:
    @pytest.mark.asyncio
    async def test_401_triggers_refresh(self, token_provider):
        client = GraphClient(token_provider)

        class FakeResponse:
            status = 200
            headers = {}
            async def __aenter__(self):
                return self
            async def __aexit__(self, *args):
                pass
            async def json(self):
                return {"value": "ok"}
            def raise_for_status(self):
                pass

        class Fake401:
            status = 401
            headers = {}
            async def __aenter__(self):
                return self
            async def __aexit__(self, *args):
                pass

        call_count = [0]

        class FakeSession:
            async def __aenter__(self):
                return self
            async def __aexit__(self, *args):
                pass

            def request(self, *args, **kwargs):
                call_count[0] += 1
                if call_count[0] == 1:
                    return Fake401()
                return FakeResponse()

        client._session = FakeSession()
        result = await client.get("/test")
        assert result == {"value": "ok"}
        assert token_provider._refresh_if_needed.called
