from pathlib import Path
import sys

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_DIR = PROJECT_ROOT / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

import pytest

from read_email import (
    EmailContent,
    extract_mfa_code,
    get_latest_mfa_code_from_outlook,
    get_latest_email_from_outlook,
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


def test_extract_mfa_code_finds_first_match():
    body = "Your login code is 123456. It expires soon."

    assert extract_mfa_code(body) == "123456"


def test_extract_mfa_code_respects_custom_pattern():
    body = "Use code MFA-9999 for access"

    assert extract_mfa_code(body, pattern=r"MFA-(\d{4})") == "MFA-9999"


class _FakeItems:
    def __init__(self, messages):
        self._messages = list(messages)

    def Sort(self, field, descending):
        reverse = bool(descending)
        if field == "[ReceivedTime]":
            self._messages.sort(key=lambda msg: msg.ReceivedTime, reverse=reverse)

    def GetFirst(self):
        return self._messages[0] if self._messages else None


class _FolderCollection(dict):
    def __getitem__(self, item):
        if item not in self:
            raise KeyError(item)
        return dict.__getitem__(self, item)


class _FakeFolder:
    def __init__(self, name, folders=None, items=None):
        self.Name = name
        self.Folders = _FolderCollection(folders or {})
        self.Items = items or _FakeItems([])


class _FakeNamespace:
    def __init__(self, inbox):
        self._inbox = inbox

    def GetDefaultFolder(self, index):
        assert index == 6
        return self._inbox


class _FakeDispatch:
    def __init__(self, namespace):
        self._namespace = namespace

    def GetNamespace(self, name):
        assert name == "MAPI"
        return self._namespace


def test_get_latest_email_from_outlook_returns_expected_message(monkeypatch):
    import read_email as module

    class _Message:
        def __init__(self, subject, sender, to, received_time, body):
            self.Subject = subject
            self.SenderEmailAddress = sender
            self.To = to
            self.ReceivedTime = received_time
            self.Body = body

    message = _Message(
        subject="MFA Request",
        sender="it@example.com",
        to="user@example.com",
        received_time="2024-01-01T12:00:00",
        body="Your code: 654321",
    )

    variable_folder = _FakeFolder("variable 1", items=_FakeItems([message]))
    inbox = _FakeFolder("Inbox", folders={"variable 1": variable_folder})
    namespace = _FakeNamespace(inbox)

    class _FakeWin32Client:
        def Dispatch(self, name):
            assert name == "Outlook.Application"
            return _FakeDispatch(namespace)

    monkeypatch.setattr(module, "_win32_client", _FakeWin32Client())

    email = get_latest_email_from_outlook("variable 1")

    assert isinstance(email, EmailContent)
    assert email.subject == "MFA Request"
    assert email.sender == "it@example.com"
    assert email.recipients == ("user@example.com",)
    assert email.body == "Your code: 654321"
    assert extract_mfa_code(email.body) == "654321"


def test_get_latest_mfa_code_from_outlook_returns_code(monkeypatch):
    import read_email as module

    class _Message:
        def __init__(self, subject, body):
            self.Subject = subject
            self.SenderEmailAddress = "it@example.com"
            self.To = "user@example.com"
            self.ReceivedTime = "2024-01-01T12:00:00"
            self.Body = body

    message = _Message(subject="MFA", body="Security code: 112233")

    folder = _FakeFolder("variable 1", items=_FakeItems([message]))
    inbox = _FakeFolder("Inbox", folders={"variable 1": folder})
    namespace = _FakeNamespace(inbox)

    class _FakeWin32Client:
        def Dispatch(self, name):
            assert name == "Outlook.Application"
            return _FakeDispatch(namespace)

    monkeypatch.setattr(module, "_win32_client", _FakeWin32Client())

    code = get_latest_mfa_code_from_outlook("variable 1")

    assert code == "112233"


def test_get_latest_mfa_code_from_outlook_raises_when_missing(monkeypatch):
    import read_email as module

    class _Message:
        def __init__(self):
            self.Subject = "No code"
            self.SenderEmailAddress = "it@example.com"
            self.To = "user@example.com"
            self.ReceivedTime = "2024-01-01T12:00:00"
            self.Body = "Hello"

    folder = _FakeFolder("Inbox", items=_FakeItems([_Message()]))
    namespace = _FakeNamespace(folder)

    class _FakeWin32Client:
        def Dispatch(self, name):
            assert name == "Outlook.Application"
            return _FakeDispatch(namespace)

    monkeypatch.setattr(module, "_win32_client", _FakeWin32Client())

    with pytest.raises(LookupError):
        get_latest_mfa_code_from_outlook("Inbox")


def test_get_latest_email_from_outlook_raises_when_missing_dependency(monkeypatch):
    import read_email as module

    monkeypatch.setattr(module, "_win32_client", None)

    with pytest.raises(RuntimeError):
        get_latest_email_from_outlook("Inbox")
