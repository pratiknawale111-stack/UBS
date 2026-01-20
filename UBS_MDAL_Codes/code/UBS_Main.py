from gmail_reader import GmailDownloader
from gmail_sender import GmailSender
from sap_automation import run_sap_upload
from sap_killer import close_sap
from datetime import datetime
import sys

# =========================================
# WEEKEND CHECK (0=Mon, 6=Sun)
# =========================================
if datetime.today().weekday() >= 5:
    print("🚫 Weekend detected. Bot will not run.")
    sys.exit(0)

# =========================================
# EMAIL CONFIG
# =========================================
USER_EMAIL = "anuj.gautam.ext@holcim.com"
DEV_EMAIL = "anuj.gautam.ext@holcim.com"

sender = GmailSender(
    credentials_path=r"D:\UBS MDAL\credentials_gmail.json",
    token_path=r"D:\UBS MDAL\token_sender.json"
)

# =========================================
# EMAIL HELPERS
# =========================================
def send_success():
    body = (
        "Dear Team,\n\n"
        "The UBS MDAL bot completed successfully.\n"
        "Please refer to the log file for per-file upload status.\n\n"
        "Regards,\n"
        "UBS-MDAL Bot"
    )

    sender.send_email(USER_EMAIL, "UBS MDAL Bot – SUCCESS", body)
    sender.send_email(DEV_EMAIL, "UBS MDAL Bot – SUCCESS", body)


def send_failure(reason):
    body = (
        "Dear Team,\n\n"
        "The UBS MDAL bot has FAILED.\n\n"
        "Failure Details:\n"
        f"{reason}\n\n"
        "Please check the bot log and screenshots for more details.\n\n"
        "Regards,\n"
        "UBS-MDAL Bot"
    )

    sender.send_email(USER_EMAIL, "UBS MDAL Bot – FAILED", body)
    sender.send_email(DEV_EMAIL, "UBS MDAL Bot – FAILED", body)

# =========================================
# MAIN EXECUTION
# =========================================
try:
    print("📥 Downloading invoices from Gmail...")

    reader = GmailDownloader(
        credentials_path=r"D:\UBS MDAL\credentials_gmail.json",
        token_path=r"D:\UBS MDAL\token_reader.json",
        download_folder=r"D:\UBS MDAL\Invoices",
        query='subject:"Extras MICB." has:attachment'
    )
    reader.download_latest()

    print("📤 Uploading invoices to SAP...")

    sap_success, sap_message = run_sap_upload(r"D:\UBS MDAL\creds.txt")

    if sap_success:
        send_success()
    else:
        send_failure(sap_message)

except Exception as e:
    send_failure(f"Controller error: {str(e)}")

finally:
    close_sap()
