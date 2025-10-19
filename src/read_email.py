"""Utility functions for parsing .eml email files."""
from __future__ import annotations

from dataclasses import dataclass
from email import policy
from email.parser import BytesParser
from pathlib import Path
from typing import Iterable, Optional


@dataclass
class EmailContent:
    """Structured representation of email metadata and body."""

    subject: Optional[str]
    sender: Optional[str]
    recipients: tuple[str, ...]
    date: Optional[str]
    body: str


def _extract_body(message) -> str:
    """Extract the most relevant text body from an email message."""
    if message.is_multipart():
        # Prefer the first text/plain part.
        for part in message.walk():
            if part.get_content_type() == "text/plain":
                payload = part.get_payload(decode=True)
                if payload is not None:
                    return payload.decode(part.get_content_charset() or "utf-8", errors="replace")
        return ""

    payload = message.get_payload(decode=True)
    if payload is None:
        return ""
    return payload.decode(message.get_content_charset() or "utf-8", errors="replace")


def parse_email(path: str | Path) -> EmailContent:
    """Parse a .eml file and return structured email content."""
    path = Path(path)
    with path.open("rb") as fp:
        message = BytesParser(policy=policy.default).parse(fp)

    return EmailContent(
        subject=message.get("subject"),
        sender=message.get("from"),
        recipients=tuple(message.get_all("to", [])),
        date=message.get("date"),
        body=_extract_body(message),
    )


def iter_email_bodies(paths: Iterable[str | Path]) -> Iterable[str]:
    """Yield body text for each email provided."""
    for path in paths:
        yield parse_email(path).body


__all__ = ["EmailContent", "parse_email", "iter_email_bodies"]
