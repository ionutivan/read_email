"""Email parsing utilities."""

from .read_email import EmailContent, iter_email_bodies, parse_email

__all__ = ["EmailContent", "iter_email_bodies", "parse_email"]
