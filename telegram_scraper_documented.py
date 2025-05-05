import logging
import gspread
import pandas as pd
from datetime import datetime, timezone, timedelta
from telethon.sync import TelegramClient
from oauth2client.service_account import ServiceAccountCredentials
from smtplib import SMTP
from email.mime.text import MIMEText


class TelegramChannelScraper:
    """
    Scrapes messages from multiple Telegram channels and saves each channel's messages
    into a separate worksheet in a Google Sheet.
    """

    def __init__(
        self, username, phone, api_id, api_hash, xlsx_path, sheet_name, creds_file
    ):
        self.user_name = username
        self.phone = phone
        self.api_id = api_id
        self.api_hash = api_hash
        self.xlsx_path = xlsx_path
        self.sheet_name = sheet_name
        self.creds_file = creds_file
        self.client = TelegramClient(self.user_name, api_id, api_hash)
        self.sheet = None
        self.now = datetime.now(timezone.utc)
        self.days_ago = self.now - timedelta(days=10)
        self._setup_logging()

    def _setup_logging(self):
        logging.basicConfig(
            level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
        )

    def connect_google_sheet(self):
        """
        Connects to Google Sheets using service account credentials.
        """
        scope = [
            "https://spreadsheets.google.com/feeds",
            "https://www.googleapis.com/auth/drive",
        ]
        creds = ServiceAccountCredentials.from_json_keyfile_name(self.creds_file, scope)
        client = gspread.authorize(creds)
        self.sheet = client.open(self.sheet_name)
        logging.info("Connected to Google Sheet: %s", self.sheet_name)

    def get_channels_from_excel(self):
        """
        Reads Telegram channel URLs from an Excel file.
        """
        df = pd.read_excel(self.xlsx_path, sheet_name=0)
        channels = df.iloc[:, 0].dropna().tolist()
        logging.info("Loaded %d channels from Excel", len(channels))
        return channels

    def scrape_channels(self):
        """
        Scrapes messages from each channel and stores them in separate sheets.
        """
        channels = self.get_channels_from_excel()
        with self.client:
            for channel in channels:
                try:
                    messages = []
                    for message in self.client.iter_messages(channel):
                        if self.days_ago <= message.date <= self.now:
                            if message.text:
                                messages.append(
                                    [
                                        message.text,
                                        str(message.date.strftime("%Y-%m-%d %H:%M:%S")),
                                    ]
                                )
                    sheet_title = message.chat.title
                    try:
                        ws = self.sheet.worksheet(sheet_title)
                        self.sheet.del_worksheet(ws)
                    except:
                        pass
                    ws = self.sheet.add_worksheet(
                        title=sheet_title, rows="1000", cols="2"
                    )

                    ws.append_row(["Message", "Timestamp"])
                    for row in messages:
                        ws.append_row(row)
                    logging.info("Scraped and saved messages for channel: %s", channel)
                except Exception as e:
                    logging.error("Failed to scrape channel %s: %s", channel, str(e))

    def send_notification(self, creds_excel_path):
        """
        Sends an email notification using credentials from an Excel file.
        """
        df = pd.read_excel(creds_excel_path)
        username = df.loc[0, "Sender"]
        app_password = df.loc[0, "App password"]
        to_email = df.loc[0, "Receiver"]

        msg = MIMEText("Telegram scraping task is completed successfully.")
        msg["Subject"] = "Telegram Scraper Notification"
        msg["From"] = username
        msg["To"] = to_email

        try:
            with SMTP("smtp.gmail.com", 587) as smtp:
                smtp.starttls()
                smtp.login(username, app_password)
                smtp.sendmail(username, [to_email], msg.as_string())
                logging.info("Notification email sent to %s", to_email)
        except Exception as e:
            logging.error("Failed to send email: %s", str(e))


if __name__ == "__main__":
    scraper = TelegramChannelScraper(
        username="your_username",
        phone="+00000000000",
        api_id="YOUR_API_ID",  # Replace with your Telegram API ID
        api_hash="YOUR_API_HASH",  # Replace with your Telegram API Hash
        xlsx_path="your_channels.xlsx",  # Replace with the path to your Excel file
        sheet_name="YourGoogleSheet",  # Replace with your Google Sheet name
        creds_file="your_creds.json",  # Replace with your Google credentials JSON
    )
    scraper.connect_google_sheet()
    scraper.scrape_channels()
    scraper.send_notification(
        "email_credentials.xlsx"
    )  # Replace with your Excel containing email creds
