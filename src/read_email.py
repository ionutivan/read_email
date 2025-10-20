"""Utility functions for parsing Outlook messages and .eml email files."""
from __future__ import annotations

import argparse
import re
from dataclasses import dataclass
from email import policy
from email.parser import BytesParser
from pathlib import Path
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


try:  # pragma: no cover - optional dependency
    import win32com.client as _win32_client
except ImportError:  # pragma: no cover - optional dependency
    _win32_client = None


def _split_addresses(addresses: Optional[str]) -> tuple[str, ...]:
    if not addresses:
        return ()
    parts = [addr.strip() for addr in re.split(r"[;,]", addresses) if addr.strip()]
    return tuple(parts)


def _resolve_outlook_folder(namespace, folder_path: str):
    folder = namespace.GetDefaultFolder(6)  # 6 == inbox
    path = [part.strip() for part in folder_path.split("/") if part.strip()]
    if path and path[0].lower() == getattr(folder, "Name", "Inbox").lower():
        path = path[1:]

    for part in path:
        try:
            folder = folder.Folders[part]
        except Exception as exc:  # pragma: no cover - depends on COM runtime
            raise LookupError(
                f"Unable to locate Outlook folder segment '{part}' in path '{folder_path}'."
            ) from exc
    return folder


def get_latest_email_from_outlook(folder_path: str = "Inbox") -> EmailContent:
    """Return the most recent email from the specified Outlook folder."""

    if _win32_client is None:  # pragma: no cover - requires Outlook runtime
        raise RuntimeError(
            "win32com is not available. Reading directly from Outlook requires pywin32 on Windows."
        )

    outlook = _win32_client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    folder = _resolve_outlook_folder(namespace, folder_path)

    items = folder.Items
    items.Sort("[ReceivedTime]", True)
    message = items.GetFirst()
    if message is None:
        raise LookupError(f"No messages found in Outlook folder '{folder_path}'.")

    return EmailContent(
        subject=getattr(message, "Subject", None),
        sender=getattr(message, "SenderEmailAddress", None),
        recipients=_split_addresses(getattr(message, "To", None)),
        date=str(getattr(message, "ReceivedTime", None)) if hasattr(message, "ReceivedTime") else None,
        body=getattr(message, "Body", "") or "",
    )


def extract_mfa_code(body: str, pattern: str = r"\b\d{6}\b") -> Optional[str]:
    """Extract the first MFA code from the email body using the provided pattern."""

    match = re.search(pattern, body)
    if match:
        return match.group(0)
    return None


def get_latest_mfa_code_from_outlook(
    folder_path: str = "Inbox", pattern: str = r"\b\d{6}\b"
) -> str:
    """Return the first MFA code from the latest email in the Outlook folder.

    Parameters
    ----------
    folder_path:
        Path to the folder within the default Outlook mailbox (e.g. "Inbox/variable 1").
    pattern:
        Regular expression pattern describing the MFA code format. Defaults to a six digit code.

    Raises
    ------
    LookupError
        If the Outlook folder contains no messages or the latest message does not contain
        a matching code.
    RuntimeError
        If ``win32com`` is unavailable on the current platform.
    """

    email = get_latest_email_from_outlook(folder_path)
    code = extract_mfa_code(email.body, pattern=pattern)
    if code is None:
        raise LookupError(
            "The latest Outlook email did not contain a value that matches the MFA code pattern."
        )
    return code


def _build_cli_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Retrieve the latest MFA code from an Outlook folder.",
    )
    parser.add_argument(
        "--folder",
        "-f",
        default="Inbox",
        help=(
            "Path to the Outlook folder to inspect, relative to the default inbox. "
            "Use forward slashes to separate nested folders, for example 'Inbox/variable 1'."
        ),
    )
    parser.add_argument(
        "--pattern",
        "-p",
        default=r"\b\d{6}\b",
        help="Regular expression describing the MFA code format.",
    )
    return parser


def main(argv: Optional[Sequence[str]] = None) -> int:
    """Console script that prints the latest MFA code from Outlook."""

    parser = _build_cli_parser()
    args = parser.parse_args(argv)

    try:
        code = get_latest_mfa_code_from_outlook(args.folder, pattern=args.pattern)
    except Exception as exc:  # pragma: no cover - exercised via tests
        parser.exit(1, f"Error: {exc}\n")

    print(code)
    return 0


__all__ = [
    "EmailContent",
    "parse_email",
    "iter_email_bodies",
    "get_latest_email_from_outlook",
    "get_latest_mfa_code_from_outlook",
    "extract_mfa_code",
]


if __name__ == "__main__":  # pragma: no cover - exercised via CLI invocation
    raise SystemExit(main())
