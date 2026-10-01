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

### 1. Download the project

With Git installed, run:

```bash
git clone https://github.com/jonrugova/gmail-retail-outreach-automation.git
cd gmail-retail-outreach-automation
```

Alternatively, select **Code → Download ZIP** on GitHub, extract the ZIP, and open a terminal in the extracted project folder.

Run all remaining commands from the project folder so the script can find its credentials, tracker, and log files.

### 2. Set up Python and dependencies

The project was tested locally with **Python 3.13 on an Intel Mac**. Python 3.13 is recommended to reproduce that setup. The pinned `cryptography` version in `requirements.txt` is retained from that working installation.

On macOS or Linux with Python 3.13 installed:

```bash
python3.13 -m venv .venv
source .venv/bin/activate
python --version
python -m pip install -r requirements.txt
```

On Windows with Python 3.13 installed:

```powershell
py -3.13 -m venv .venv
.venv\Scripts\Activate.ps1
python --version
python -m pip install -r requirements.txt
```

Confirm that `python --version` reports Python 3.13 before continuing.

### 3. Configure Gmail OAuth

Enable the Gmail API in a Google Cloud project, configure its OAuth consent screen, and create an OAuth client with application type **Desktop app**. If the app is in testing mode, add the Gmail account you intend to use as a test user.

Download the OAuth client credentials and save them in the project folder as:

```text
credentials.json
```

The script requests the `gmail.modify` scope. On the first run, it opens a browser for authorization and saves `token.json` locally. Authorize the Gmail account containing the historical inquiries.

For Google Cloud setup details, see the [official Gmail API Python quickstart](https://developers.google.com/workspace/gmail/api/quickstart/python).

### 4. Prepare the Excel tracker

Place an Excel file in the project folder named:

```text
retail_outreach_tracker.xlsx
```

Use a worksheet named `Retail Outreach` with column headers in the first row:

| Email | Email Sent | Client Replied |
| --- | --- | --- |
| retailer@example.com | | |

The script requires `Email` and `Email Sent`. `Client Replied` is an optional column for manual tracking; the script does not update it.

Rows marked `Yes` under `Email Sent` are skipped. After a successful send, the script marks the row `Yes` and saves the workbook. Close the tracker in Excel before running the script so it can be saved.

### 5. Customize the local campaign

The repository includes generic example settings. Before using it for your own campaign, edit the **CONFIGURATION** section near the top of your local `reply_retailers.py`:

| Setting | What to change |
| --- | --- |
| `EXCEL_FILE` / `SHEET_NAME` | Match your tracker filename and worksheet if you use different names. |
| `OUR_EMAILS` | Replace `team@example.com` with your internal email addresses and aliases, using lowercase addresses. These help exclude internal messages when selecting an inbound message. |
| `TEMPLATE` | Replace the sample partnership email with your campaign wording and signature. Keep the `{greeting}` placeholder for personalization. |
| `MARKER_TEXT` | Use a distinctive sentence that appears exactly in `TEMPLATE`. The script searches sent messages in the selected thread for this text to detect prior campaign replies. Keep it unchanged when resuming the same campaign. |
| `MAX_CONTACTS_PER_RUN` / `MAX_SEND_PER_RUN` | Adjust the batch limits if needed. Both currently default to 30. |

`OUR_EMAILS` does not select the sending account. Emails are sent using the Gmail account authorized in step 3.

Keep company-specific addresses and campaign wording local. Because `reply_retailers.py` is tracked by Git, its edits are not protected by `.gitignore`; do not commit private campaign settings to a public repository.

### 6. Test, then enable sending

The published script defaults to:

```python
DRY_RUN = True
```

Run:

```bash
python reply_retailers.py
```

A dry run finds candidate threads and writes results to `retail_reply_log.csv`, but sends no emails and does not update the tracker.

Review the matched threads and subjects before sending. Once your local settings and matches are checked, change `DRY_RUN = False` in your local script and run the same command again.

## Privacy

Real customer data, OAuth credentials, Gmail tokens, campaign logs, and production Excel files are intentionally excluded from this repository.

## Use Case

Originally designed as an internal workflow automation for managing a large backlog of historical retailer and distributor inquiries while preserving the original Gmail conversation context.
