"""Tests for teams_graph auth module — token discovery & refresh."""

import os
import sys
import pytest

# Add plugin path
sys.path.insert(0, os.path.expanduser("~/.hermes/hermes-agent"))
sys.path.insert(0, os.path.expanduser("~/.hermes/hermes-agent/plugins/platforms/teams_graph"))

from unittest.mock import patch, MagicMock
from pathlib import Path


@pytest.fixture(autouse=True)
def clear_env():
    """Remove auth-related env vars before each test."""
    for key in list(os.environ):
        if any(k in key for k in ("TEAMS_GRAPH", "M365_ACCESS", "M365_REFRESH", "M365_TOKEN_PASSWORD", "AZURE_CLIENT_ID", "AZURE_TENANT_ID")):
            del os.environ[key]
    yield


class TestGraphToken:
    def test_not_expired_without_timestamp(self):
        from auth import GraphToken
        t = GraphToken(access_token="x")
        assert not t.is_expired

    def test_not_expired_when_fresh(self):
        import time
        from auth import GraphToken
        t = GraphToken(access_token="x", expires_at=time.time() + 3600)
        assert not t.is_expired

    def test_expired_within_buffer(self):
        import time
        from auth import GraphToken
        t = GraphToken(access_token="x", expires_at=time.time() - 200)
        assert t.is_expired


class TestTokenProviderFromEnv:
    def test_explicit_teams_graph_token_first(self):
        from auth import TokenProvider
        os.environ["TEAMS_GRAPH_ACCESS_TOKEN"] = "tg-token"
        os.environ["M365_ACCESS_TOKEN"] = "m365-token"

        tp = TokenProvider()
        token = tp._from_env_var("TEAMS_GRAPH_ACCESS_TOKEN", "TEAMS_GRAPH_REFRESH_TOKEN")
        assert token is not None
        assert token.access_token == "tg-token"

    def test_m365_token_second(self):
        from auth import TokenProvider
        os.environ["M365_ACCESS_TOKEN"] = "m365-token"

        tp = TokenProvider()
        token = tp._from_env_var("M365_ACCESS_TOKEN", "M365_REFRESH_TOKEN")
        assert token is not None
        assert token.access_token == "m365-token"

    def test_returns_none_when_no_token(self):
        from auth import TokenProvider
        tp = TokenProvider()
        token = tp._from_env_var("TEAMS_GRAPH_ACCESS_TOKEN", "TEAMS_GRAPH_REFRESH_TOKEN")
        assert token is None

    def test_includes_refresh_token(self):
        from auth import TokenProvider
        os.environ["TEAMS_GRAPH_ACCESS_TOKEN"] = "at"
        os.environ["TEAMS_GRAPH_REFRESH_TOKEN"] = "rt"

        tp = TokenProvider()
        token = tp._from_env_var("TEAMS_GRAPH_ACCESS_TOKEN", "TEAMS_GRAPH_REFRESH_TOKEN")
        assert token.refresh_token == "rt"


class TestTokenProviderDefaults:
    def test_default_client_id(self):
        from auth import TokenProvider
        tp = TokenProvider()
        assert tp.client_id == "cba17ea1-24d3-4159-85d5-237430e4bd6c"

    def test_default_tenant_id(self):
        from auth import TokenProvider
        tp = TokenProvider()
        assert tp.tenant_id == "c80e6e0a-e825-4c48-bbcd-8b580f0090f9"

    def test_env_overrides_defaults(self):
        from auth import TokenProvider
        os.environ["TEAMS_GRAPH_CLIENT_ID"] = "my-client"
        os.environ["TEAMS_GRAPH_TENANT_ID"] = "my-tenant"
        tp = TokenProvider()
        assert tp.client_id == "my-client"
        assert tp.tenant_id == "my-tenant"


class TestTokenProviderM365SkillTokens:
    def test_returns_none_when_no_tokens_file(self):
        from auth import TokenProvider
        tp = TokenProvider()
        # Mock path to non-existent file
        with patch.object(Path, 'exists', return_value=False):
            import asyncio
            result = asyncio.new_event_loop().run_until_complete(tp._from_m365_skill_tokens())
            assert result is None


class TestTokenCaching:
    def test_get_token_returns_cached_when_valid(self):
        from auth import TokenProvider, GraphToken
        import time

        tp = TokenProvider()
        tp._current_token = GraphToken(
            access_token="cached",
            expires_at=time.time() + 3600,
        )
        import asyncio
        result = asyncio.new_event_loop().run_until_complete(tp.get_token())
        assert result.access_token == "cached"