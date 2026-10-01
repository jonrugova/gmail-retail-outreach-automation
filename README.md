# Gmail Retail Outreach Automation

A production-oriented Python automation for re-engaging historical retail and distribution leads through Gmail.

The system reads outreach status from an Excel tracker, locates the correct historical Gmail conversation, sends a personalized reply inside the existing thread, prevents duplicate outreach, and updates campaign status only after Gmail confirms a successful send.

## Features

- Gmail API integration with OAuth 2.0
- Replies inside the original Gmail thread
- Excel-based outreach tracking
- Automatic skip for previously contacted leads
- Duplicate-send protection
- Shopify contact-form email detection
- Retail/wholesale inquiry scoring
- Personalized greetings when a name can be identified
- Gmail API rate-limit handling with automatic backoff
- Dry-run mode for safe testing
- Automatic Excel updates after confirmed sends
- CSV activity logging
- Configurable batch limits

## Tech Stack

- Python
- Gmail API
- Google OAuth 2.0
- OpenPyXL
- Google API Python Client

## How It Works

1. Reads leads from an Excel outreach tracker.
2. Skips contacts already marked as sent.
3. Searches Gmail for the contact's historical messages.
4. Scores candidate emails for retail-partnership relevance.
5. Selects the most relevant Gmail thread.
6. Checks whether the outreach campaign was already sent.
7. Identifies the latest inbound message.
8. Creates a threaded reply using Gmail message headers.
9. Sends through the Gmail API.
10. Marks the lead as sent only after Gmail confirms delivery.

## Safety Features

The project includes several safeguards for production use:

- `DRY_RUN` mode
- Maximum contacts per run
- Maximum sends per run
- Rate-limit retry/backoff
- Duplicate campaign detection
- Immediate Excel persistence after each successful send
- Credentials and customer data excluded through `.gitignore`

## Setup

Create a virtual environment:

```bash
python3 -m venv .venv
source .venv/bin/activate
```

Install dependencies:

```bash
pip install -r requirements.txt
```

Create a Google Cloud OAuth Desktop application with Gmail API access and save the downloaded OAuth credentials as:

```text
credentials.json
```

Prepare an Excel file named:

```text
retail_outreach_tracker.xlsx
```

with a worksheet named:

```text
Retail Outreach
```

and columns including:

```text
Email
Email Sent
Client Replied
```

Start with:

```python
DRY_RUN = True
```

Then run:

```bash
python reply_retailers.py
```

## Privacy

Real customer data, OAuth credentials, Gmail tokens, campaign logs, and production Excel files are intentionally excluded from this repository.

## Use Case

Originally designed as an internal workflow automation for managing a large backlog of historical retailer and distributor inquiries while preserving the original Gmail conversation context.
