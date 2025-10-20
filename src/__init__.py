"""Email parsing utilities."""

from .read_email import (
    EmailContent,
    extract_mfa_code,
    get_latest_email_from_outlook,
    get_latest_mfa_code_from_outlook,
    iter_email_bodies,
    main,
    parse_email,
)

__all__ = [
    "EmailContent",
    "extract_mfa_code",
    "get_latest_email_from_outlook",
    "get_latest_mfa_code_from_outlook",
    "iter_email_bodies",
    "main",
    "parse_email",
]
