# read_email

Utilities for parsing `.eml` email files.

The primary entry point is `parse_email`, which returns an `EmailContent`
dataclass containing:

- `subject`: the email subject line
- `sender`: the raw `From` header value
- `recipients`: a tuple of values extracted from the `To` header
- `date`: the timestamp string from the `Date` header
- `body`: the decoded plain-text body content

You can also iterate over multiple files with `iter_email_bodies` to obtain
only the message bodies.

## Usage

```python
from src.read_email import parse_email

email = parse_email("tests/data/sample_email.eml")
print(email.subject)
print(email.body)
```

## Tests

Run unit tests with `pytest`:

```bash
pytest
```
