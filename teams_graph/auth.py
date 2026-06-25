"""Token discovery & refresh for Teams Graph integration.

Priority:
1. TEAMS_GRAPH_ACCESS_TOKEN env var
2. M365_ACCESS_TOKEN + M365_REFRESH_TOKEN env vars
3. M365 skill encrypted tokens (~/.hermes/skills/m365/config/tokens.enc)
4. Device-code OAuth2 flow (MSAL)
"""

import os
import json
import time
import logging
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

logger = logging.getLogger(__name__)

DEFAULT_TENANT = "c80e6e0a-e825-4c48-bbcd-8b580f0090f9"
DEFAULT_CLIENT = "cba17ea1-24d3-4159-85d5-237430e4bd6c"
DEFAULT_SCOPES = ["https://graph.microsoft.com/.default"]


@dataclass
class GraphToken:
    access_token: str
    refresh_token: Optional[str] = None
    expires_at: Optional[float] = None
    tenant_id: Optional[str] = None
    client_id: Optional[str] = None

    @property
    def is_expired(self) -> bool:
        if self.expires_at is None:
            return False
        return time.time() > self.expires_at - 300


class TokenProvider:
    """Discovers and refreshes M365 Graph API tokens."""

    def __init__(
        self,
        client_id: Optional[str] = None,
        tenant_id: Optional[str] = None,
        client_secret: Optional[str] = None,
        scopes: Optional[list[str]] = None,
    ):
        self.client_id = (
            client_id
            or os.getenv("TEAMS_GRAPH_CLIENT_ID")
            or os.getenv("M365_CLIENT_ID")
            or os.getenv("AZURE_CLIENT_ID")
            or DEFAULT_CLIENT
        )
        self.tenant_id = (
            tenant_id
            or os.getenv("TEAMS_GRAPH_TENANT_ID")
            or os.getenv("M365_TENANT_ID")
            or os.getenv("AZURE_TENANT_ID")
            or DEFAULT_TENANT
        )
        self.client_secret = client_secret or os.getenv("TEAMS_GRAPH_CLIENT_SECRET")
        self.scopes = scopes or DEFAULT_SCOPES
        self._current_token: Optional[GraphToken] = None

    async def get_token(self, allow_device_code: bool = False) -> GraphToken:
        """Return a valid token, refreshing if needed."""
        if self._current_token and not self._current_token.is_expired:
            return self._current_token
        token = await self._acquire_token(allow_device_code=allow_device_code)
        self._current_token = token
        return token

    async def _acquire_token(self, allow_device_code: bool = False) -> GraphToken:
        """Try all token sources in priority order."""
        # 1. Explicit override
        token = self._from_env_var("TEAMS_GRAPH_ACCESS_TOKEN", "TEAMS_GRAPH_REFRESH_TOKEN")
        if token:
            logger.info("Using TEAMS_GRAPH_ACCESS_TOKEN")
            return token

        # 2. Shared M365 env vars
        token = self._from_env_var("M365_ACCESS_TOKEN", "M365_REFRESH_TOKEN")
        if token:
            logger.info("Using shared M365 env vars")
            return token

        # 3. M365 skill encrypted tokens
        token = await self._from_m365_skill_tokens()
        if token:
            # If the stored token is stale but we have a refresh token, try refreshing
            if token.is_expired and token.refresh_token:
                logger.info("M365 skill token expired — attempting refresh via MSAL")
                self._current_token = token  # stash for _refresh_if_needed
                refreshed = await self._refresh_if_needed()
                if refreshed:
                    logger.info("M365 token refreshed successfully")
                    return self._current_token
                logger.warning("M365 token refresh FAILED — check refresh token validity")
            elif not token.is_expired:
                logger.info("Using M365 skill tokens")
                return token

        # 4. Device-code fallback (only when allowed — skip on gateway startup)
        if not allow_device_code:
            raise RuntimeError(
                "No token available from env vars or M365 skill tokens. "
                "Set TEAMS_GRAPH_ACCESS_TOKEN or run device-code auth separately."
            )
        logger.info("No existing tokens found — starting device-code flow")
        return await self._device_code_flow()

    def _from_env_var(self, access_key: str, refresh_key: str) -> Optional[GraphToken]:
        access = os.getenv(access_key, "").strip()
        if not access:
            return None
        token = GraphToken(
            access_token=access,
            refresh_token=os.getenv(refresh_key, "").strip() or None,
            tenant_id=self.tenant_id,
            client_id=self.client_id,
        )
        return token

    async def _from_m365_skill_tokens(self) -> Optional[GraphToken]:
        """Decrypt M365 skill tokens.enc if available."""
        tokens_path = Path.home() / ".hermes" / "skills" / "m365" / "config" / "tokens.enc"
        if not tokens_path.exists():
            logger.info("M365 skill tokens not found at %s", tokens_path)
            return None
        try:
            from cryptography.fernet import Fernet
            from cryptography.hazmat.primitives import hashes
            from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC as PBKDF2
            import base64

            password = os.getenv("M365_TOKEN_PASSWORD")
            if password:
                salt = tokens_path.with_suffix(".salt")
                if salt.exists():
                    salt_bytes = salt.read_bytes()
                else:
                    salt_bytes = b"hermes-m365-skill"
                kdf = PBKDF2(
                    algorithm=hashes.SHA256(),
                    length=32,
                    salt=salt_bytes,
                    iterations=100_000,
                )
                key = base64.urlsafe_b64encode(kdf.derive(password.encode()))
            else:
                key_path = tokens_path.parent / "token.key"
                if not key_path.exists():
                    logger.info("M365 token.key not found at %s", key_path)
                    return None
                key = key_path.read_bytes()

            fernet = Fernet(key)
            decrypted = fernet.decrypt(tokens_path.read_bytes())
            data = json.loads(decrypted)
            logger.info("Decrypted M365 skill tokens successfully (has_refresh=%s)",
                        "refresh_token" in data)
            # We don't know when the token was issued, so treat as expired and
            # rely on the refresh token in _acquire_token() to obtain a valid one.
            return GraphToken(
                access_token=data.get("access_token", ""),
                refresh_token=data.get("refresh_token"),
                expires_at=0,  # force refresh — we don't know issue time
                tenant_id=self.tenant_id,
                client_id=self.client_id,
            )
        except Exception as e:
            logger.error("Could not decrypt M365 skill tokens: %s (path=%s, key_exists=%s)",
                         e, tokens_path, tokens_path.parent.joinpath("token.key").exists())
            return None

    async def _device_code_flow(self) -> GraphToken:
        """MSAL device-code OAuth2 flow."""
        import msal

        app = msal.PublicClientApplication(
            client_id=self.client_id,
            authority=f"https://login.microsoftonline.com/{self.tenant_id}",
        )
        flow = app.initiate_device_flow(scopes=self.scopes)
        if "user_code" not in flow:
            raise RuntimeError("Device code flow initiation failed")

        print(f"\n  🔑 Open: {flow['verification_uri']}")
        print(f"  📟 Enter code: {flow['user_code']}\n")

        result = app.acquire_token_by_device_flow(flow)
        if "access_token" not in result:
            raise RuntimeError(
                f"Auth failed: {result.get('error_description', result.get('error', 'unknown'))}"
            )

        return GraphToken(
            access_token=result["access_token"],
            refresh_token=result.get("refresh_token"),
            expires_at=time.time() + result.get("expires_in", 3600),
            tenant_id=self.tenant_id,
            client_id=self.client_id,
        )

    async def _refresh_if_needed(self) -> bool:
        """Try to refresh the current token. Returns True if refresh succeeded."""
        if not self._current_token or not self._current_token.refresh_token:
            return False
        try:
            import msal
            app = msal.PublicClientApplication(
                client_id=self.client_id,
                authority=f"https://login.microsoftonline.com/{self.tenant_id}",
            )
            result = app.acquire_token_by_refresh_token(
                refresh_token=self._current_token.refresh_token,
                scopes=self.scopes,
            )
            if "access_token" in result:
                self._current_token = GraphToken(
                    access_token=result["access_token"],
                    refresh_token=result.get("refresh_token", self._current_token.refresh_token),
                    expires_at=time.time() + result.get("expires_in", 3600),
                    tenant_id=self.tenant_id,
                    client_id=self.client_id,
                )
                logger.info("Token refreshed successfully")
                return True
        except Exception as e:
            logger.warning("Token refresh failed: %s", e)
        return False
