# Chatbot & Email-to-Excel Tools

A Python repository containing two separate Streamlit tools: a streamed chat demo and an email extraction/export application.

**Technology:** Python · Streamlit · OpenAI client · IMAP/Gmail API · pandas · openpyxl

## Features

- Run the original conversational demo with an API key entered in the interface.
- Connect to an inbox through IMAP or Gmail API for email extraction.
- Filter and inspect retrieved emails and export a formatted Excel workbook.
- Keep the chat and email tools as independent entry points.

## Repository guide

| Path | Purpose |
|---|---|
| [streamlit_app.py](streamlit_app.py) | Streamed conversational demo. |
| [email_to_excel_app.py](email_to_excel_app.py) | Inbox connection, preview, and workbook download UI. |
| [email_processor.py](email_processor.py) | IMAP and Gmail API email processing. |
| [excel_exporter.py](excel_exporter.py) | Excel workbook generation. |
| [README_EMAIL_TO_EXCEL.md](README_EMAIL_TO_EXCEL.md) | Email workflow documentation. |
| [INSTALLATION_GUIDE.md](INSTALLATION_GUIDE.md) | Additional setup notes. |
| [requirements.txt](requirements.txt) | Email application dependencies. |

## Requirements and current limitations

The repository also includes `notebooke7f0cd3e02.ipynb`, an unrelated traffic severity experiment. It is not required by either Streamlit entry point. Email access, OAuth setup, and hosted model availability must be configured separately.

## Getting started

```bash
git clone https://github.com/IbrahimAbdelsattar/chatbot.git
cd chatbot
```

Use a Python virtual environment:

```bash
python -m venv .venv
```

Activate it with `source .venv/bin/activate` on macOS/Linux or `.venv\Scripts\Activate.ps1` in PowerShell.

```bash
python -m pip install -r requirements.txt
python -m streamlit run email_to_excel_app.py
```

## Chat demo

The chat entry point additionally imports the OpenAI Python client:

```bash
python -m pip install openai
python -m streamlit run streamlit_app.py
```

Enter your own API key in the interface. The current source requests `gpt-3.5-turbo`; change that model setting if the configured provider no longer offers it. Streamlit session state stores the conversation for the current session.

## Email authentication

For IMAP, provide your server address and account credentials accepted by that provider. Gmail API mode requires your own OAuth client configuration (`credentials.json`) and account consent. The code may create `token.pickle` during authorization. Keep account credentials and generated tokens private.

## License

Apache-2.0. See [LICENSE](LICENSE) for the license terms.
