"""Tests for teams_graph platform adapter."""

import os
import sys
import pytest
from unittest.mock import AsyncMock, MagicMock, patch

from gateway.config import PlatformConfig
from plugins.platforms.teams_graph.adapter import (
    TeamsGraphAdapter,
    _env_enablement,
    _apply_yaml_config,
    _is_connected,
    _standalone_send,
    check_requirements,
)


@pytest.fixture(autouse=True)
def clear_env():
    for key in list(os.environ):
        if "TEAMS_GRAPH" in key:
            del os.environ[key]
    yield


@pytest.fixture
def platform_config():
    return PlatformConfig(
        enabled=True,
        extra={
            "client_id": "test-cid",
            "tenant_id": "test-tid",
        },
    )


@pytest.fixture
def adapter(platform_config):
    return TeamsGraphAdapter(platform_config)


class TestAdapterInit:
    def test_adapter_parses_config(self, adapter):
        assert adapter._tp.client_id == "test-cid"
        assert adapter._tp.tenant_id == "test-tid"

    def test_adapter_inherits_base(self, adapter):
        from gateway.platforms.base import BasePlatformAdapter
        assert isinstance(adapter, BasePlatformAdapter)


class TestAdapterSend:
    def test_send_returns_error_when_not_connected(self, adapter):
        import asyncio
        result = asyncio.new_event_loop().run_until_complete(
            adapter.send("c1", "hello")
        )
        assert result.success is False
        assert result.error == "Not connected"

    @pytest.mark.asyncio
    async def test_send_truncates_long_messages(self, adapter):
        adapter._graph = MagicMock()
        adapter._graph.send_chat_message = AsyncMock(return_value={"id": "msg1"})

        long_msg = "x" * 30000
        result = await adapter.send("c1", long_msg)
        assert len(long_msg) > 28000
        assert result.success is True
        call_arg = adapter._graph.send_chat_message.call_args[0][1]
        assert len(call_arg) <= 28000
        assert "[message truncated]" in call_arg

    @pytest.mark.asyncio
    async def test_send_returns_send_result(self, adapter):
        adapter._graph = MagicMock()
        adapter._graph.send_chat_message = AsyncMock(return_value={"id": "msg-42"})

        result = await adapter.send("c1", "hi")
        assert result.success is True
        assert result.message_id == "msg-42"

    @pytest.mark.asyncio
    async def test_send_handles_api_error(self, adapter):
        adapter._graph = MagicMock()
        adapter._graph.send_chat_message = AsyncMock(side_effect=Exception("API error"))

        result = await adapter.send("c1", "hi")
        assert result.success is False
        assert "API error" in result.error


class TestEnvEnablement:
    def test_returns_none_when_unconfigured(self):
        result = _env_enablement()
        assert result is None

    def test_seeds_client_and_tenant(self):
        os.environ["TEAMS_GRAPH_CLIENT_ID"] = "cid"
        os.environ["TEAMS_GRAPH_TENANT_ID"] = "tid"
        result = _env_enablement()
        assert result is not None
        assert result["client_id"] == "cid"
        assert result["tenant_id"] == "tid"

    def test_seeds_home_channel(self):
        os.environ["TEAMS_GRAPH_CLIENT_ID"] = "cid"
        os.environ["TEAMS_GRAPH_TENANT_ID"] = "tid"
        os.environ["TEAMS_GRAPH_HOME_CHANNEL"] = "chat-1"
        os.environ["TEAMS_GRAPH_HOME_CHANNEL_NAME"] = "My Chat"
        result = _env_enablement()
        assert result["home_channel"]["chat_id"] == "chat-1"
        assert result["home_channel"]["name"] == "My Chat"


class TestApplyYamlConfig:
    def test_seeds_env_vars(self):
        result = _apply_yaml_config(
            {},
            {"client_id": "yaml-cid", "tenant_id": "yaml-tid"}
        )
        assert result is not None
        assert result["client_id"] == "yaml-cid"
        assert os.environ["TEAMS_GRAPH_CLIENT_ID"] == "yaml-cid"

    def test_env_precedence_over_yaml(self):
        os.environ["TEAMS_GRAPH_CLIENT_ID"] = "env-cid"
        result = _apply_yaml_config(
            {},
            {"client_id": "yaml-cid", "tenant_id": "yaml-tid"}
        )
        assert os.environ["TEAMS_GRAPH_CLIENT_ID"] == "env-cid"

    def test_home_channel_extraction(self):
        result = _apply_yaml_config(
            {},
            {
                "client_id": "cid",
                "tenant_id": "tid",
                "home_channel": {"chat_id": "hc1", "name": "Ops"},
            }
        )
        assert result["home_channel"]["chat_id"] == "hc1"
        assert os.environ["TEAMS_GRAPH_HOME_CHANNEL"] == "hc1"
        assert os.environ["TEAMS_GRAPH_HOME_CHANNEL_NAME"] == "Ops"


class TestIsConnected:
    def test_false_when_no_creds(self):
        cfg = MagicMock()
        result = _is_connected(cfg)
        assert result is False


class TestCheckRequirements:
    def test_aiohttp_and_msal_available(self):
        assert check_requirements() is True


class TestStandaloneSend:
    @pytest.mark.asyncio
    async def test_requires_credentials(self):
        cfg = MagicMock()
        cfg.extra = {}
        result = await _standalone_send(cfg, "c1", "hi")
        assert "error" in result
        assert "required" in result["error"].lower()
