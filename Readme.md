````markdown
# Telegram Channel Scraper to Google Sheets

## Overview

This tool scrapes messages from a list of Telegram channels and stores the content in a Google Spreadsheet. Each channel's messages are saved to a separate sheet with timestamps. After the scraping is complete, it sends an email notification using credentials from an Excel file.

## Features

- Read list of Telegram channel URLs from Excel
- Scrape messages from each channel for the past N days
- Save each channel’s data to its own worksheet in a Google Sheet
- Supports media content (detects presence of photos)
- Sends email notification on completion
- Logs all events with timestamps

## Requirements

- Python 3.8+
- Telegram API credentials (API ID, API Hash)
- Google Sheets API credentials (`credentials.json`)
- Excel file containing:
  - Telegram channel URLs
  - Email credentials for sending notifications

## Installation

```bash
pip install telethon pandas openpyxl gspread oauth2client
````

## Setup

1. Get your Telegram API ID and Hash from [https://my.telegram.org](https://my.telegram.org).
2. Enable Google Sheets API and download `credentials.json` for service account access.
3. Prepare Excel files:

   * `channels.xlsx`: Column A contains channel URLs.
   * `email_config.xlsx`: Contains email sender details.

## Usage

```bash
python telegram_scraper_documented.py
```

* The program will:

  1. Log in to Telegram
  2. Scrape messages from each listed channel
  3. Store each channel's messages in separate sheets
  4. Send a summary notification via email

## Project Structure

```
.
├── telegram_scraper_documented.py
├── channels.xlsx
├── email_config.xlsx
├── credentials.json
└── README.md
```

## Logging

Logs are stored in the console with timestamps. Useful for debugging and monitoring.


## Author
## 🧑‍💻 Author

- Built by [Adel Alaa](https://github.com/adel-alaa-dp)
- Linkedin: https://www.linkedin.com/in/adel-alaa/
- Fiverr: https://www.fiverr.com/s/o8j7g7g
- UpWork: https://www.upwork.com/freelancers/~01c667db97cbb336b5
