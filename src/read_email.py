"""Utility functions for parsing .eml email files."""
from __future__ import annotations

from dataclasses import dataclass
from email import policy
from email.parser import BytesParser
from pathlib import Path
import argparse
import re
import sys
from typing import Iterable, Optional, Sequence


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


def find_latest_email(directory: str | Path) -> Path:
    """Return the most recently modified .eml file within ``directory``."""

    directory = Path(directory)
    if not directory.exists():
        raise FileNotFoundError(f"Directory not found: {directory}")
    if not directory.is_dir():
        raise NotADirectoryError(f"Not a directory: {directory}")

    eml_files = sorted(
        directory.glob("*.eml"),
        key=lambda path: (path.stat().st_mtime, path.name),
    )
    if not eml_files:
        raise FileNotFoundError(f"No .eml files found in {directory}")

    return eml_files[-1]


def extract_mfa_code(body: str) -> str:
    """Return the first 6-digit code found in ``body``.

    Raises ``ValueError`` when no plausible MFA code is present.
    """

    match = re.search(r"\b(\d{6})\b", body)
    if match is None:
        raise ValueError("No MFA code found in email body")
    return match.group(1)


def main(argv: Optional[Sequence[str]] = None) -> int:
    """Command-line interface for extracting MFA codes."""

    parser = argparse.ArgumentParser(
        description="Extract the most recent MFA code from .eml files.",
    )
    parser.add_argument(
        "path",
        nargs="?",
        default=".",
        help="Directory containing .eml files or a single .eml file",
    )
    args = parser.parse_args(argv)

    target = Path(args.path)
    if not target.exists():
        print(f"No such file or directory: {target}", file=sys.stderr)
        return 1

    try:
        email_path = find_latest_email(target) if target.is_dir() else target
    except (FileNotFoundError, NotADirectoryError) as exc:
        print(str(exc), file=sys.stderr)
        return 1

    try:
        email = parse_email(email_path)
        code = extract_mfa_code(email.body)
    except ValueError as exc:
        print(str(exc), file=sys.stderr)
        return 1

    print(code)
    return 0


__all__ = [
    "EmailContent",
    "parse_email",
    "iter_email_bodies",
    "find_latest_email",
    "extract_mfa_code",
    "main",
]


if __name__ == "__main__":  # pragma: no cover - CLI entry point
    raise SystemExit(main())
