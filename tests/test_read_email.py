from pathlib import Path
import os
import subprocess
import sys

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_DIR = PROJECT_ROOT / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from read_email import (
    EmailContent,
    extract_mfa_code,
    find_latest_email,
    iter_email_bodies,
    parse_email,
)


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


def test_find_latest_email_returns_most_recent_file(tmp_path):
    old_email = tmp_path / "old.eml"
    old_email.write_text("Old email")
    new_email = tmp_path / "new.eml"
    new_email.write_text("New email")

    os.utime(old_email, (1, 1))
    os.utime(new_email, (2, 2))

    latest = find_latest_email(tmp_path)

    assert latest == new_email


def test_extract_mfa_code_returns_first_six_digit_sequence():
    body = "Your login code is 654321."

    assert extract_mfa_code(body) == "654321"


def test_cli_prints_latest_mfa_code(tmp_path):
    older_email = tmp_path / "older.eml"
    newer_email = tmp_path / "newer.eml"
    older_email.write_bytes(
        (
            "From: sender@example.com\n"
            "To: recipient@example.com\n"
            "Subject: Test\n"
            "Content-Type: text/plain; charset=utf-8\n"
            "\n"
            "Your code is 111111.\n"
        ).encode("utf-8")
    )
    newer_email.write_bytes(
        (
            "From: sender@example.com\n"
            "To: recipient@example.com\n"
            "Subject: Test\n"
            "Content-Type: text/plain; charset=utf-8\n"
            "\n"
            "Your code is 222222.\n"
        ).encode("utf-8")
    )
    os.utime(older_email, (1, 1))
    os.utime(newer_email, (2, 2))

    env = os.environ.copy()
    pythonpath = env.get("PYTHONPATH")
    env["PYTHONPATH"] = f"{SRC_DIR}:{pythonpath}" if pythonpath else str(SRC_DIR)
    result = subprocess.run(
        [sys.executable, "-m", "read_email", str(tmp_path)],
        capture_output=True,
        text=True,
        env=env,
        check=False,
    )

    assert result.returncode == 0
    assert result.stdout.strip() == "222222"
