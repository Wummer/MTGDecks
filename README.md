# MTGDecks
For hobby purposes

## Setup

```bash
pip install -r requirements.txt
```

## Usage

```bash
# Google Sheets (default) + local .xlsx
python deck_safe_collection_builder.py collection.csv --deck-dir ./decks/ -o output.xlsx

# Google Sheets only (no local file)
python deck_safe_collection_builder.py collection.csv --deck-dir ./decks/

# Local .xlsx only (no Google Sheets)
python deck_safe_collection_builder.py collection.csv --deck-dir ./decks/ -o output.xlsx --no-google
```

## Google Sheets Setup

The script automatically uploads to Google Sheets on every run. First-time setup:

1. Go to [Google Cloud Console](https://console.cloud.google.com/) and create a new project
2. Enable the **Google Sheets API** and **Google Drive API** for the project
3. Go to **APIs & Services > Credentials** and create an **OAuth 2.0 Client ID** (application type: Desktop)
4. Download the credentials JSON and save it as `~/.config/gspread/credentials.json`
5. On the first run, a browser window will open for you to authorize the app. The token is cached automatically for future runs.

Use `--sheet-name "My Sheet Name"` to customize the Google Sheets document name (default: "Deck-Safe Collection").
