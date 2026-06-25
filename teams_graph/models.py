"""Data models for the Teams Graph platform adapter."""

from dataclasses import dataclass, field
from typing import Any, Optional


@dataclass
class TeamsUser:
    id: str
    display_name: str
    email: Optional[str] = None

    @classmethod
    def from_graph(cls, data: dict[str, Any]) -> "TeamsUser":
        return cls(
            id=data.get("id", ""),
            display_name=data.get("displayName", "Unknown"),
            email=data.get("mail") or data.get("userPrincipalName"),
        )


@dataclass
class TeamsChatMessage:
    id: str
    chat_id: str
    content: str
    content_type: str = "text"
    sender: Optional[TeamsUser] = None
    created_at: Optional[str] = None
    raw: dict[str, Any] = field(default_factory=dict)

    @property
    def is_from_self(self) -> bool:
        """Check if this message was sent by our own identity."""
        return False  # determined by comparison at adapter level
