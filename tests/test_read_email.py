from pathlib import Path
import sys

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_DIR = PROJECT_ROOT / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from read_email import EmailContent, iter_email_bodies, parse_email


FIXTURE = Path(__file__).parent / "data" / "sample_email.eml"


def test_parse_email_returns_expected_content():
    email = parse_email(FIXTURE)

    assert isinstance(email, EmailContent)
    assert email.subject == "Test Email"
    assert "alice@example.com" in email.sender
    assert email.recipients == ("bob@example.com",)
    assert "This is a test email" in email.body


def test_iter_email_bodies_yields_body_text():
    bodies = list(iter_email_bodies([FIXTURE]))

    assert len(bodies) == 1
    assert "This is a test email" in bodies[0]
