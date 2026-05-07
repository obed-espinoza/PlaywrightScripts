SCRIPT_INPUTS = []

import os
import re
import sys
import time
import random
import smtplib
import traceback
import zipfile
from datetime import datetime
from email.mime.text import MIMEText
from pathlib import Path

from credential_loader import require_secret, get_secret
import pandas as pd
from playwright.sync_api import sync_playwright
from twilio.rest import Client

# ── Paths (relative to this script, not CWD) ─────────────────
SCRIPT_DIR = Path(__file__).resolve().parent
AUTH_FILE = SCRIPT_DIR / "auth_state.json"
DOWNLOAD_DIR = SCRIPT_DIR / "downloads"
VALIDATED_CSV = SCRIPT_DIR / "validated_tickets.csv"

# ── SharePoint file URL ──────────────────────────────────────
SHAREPOINT_URL = (
    "https://kbsservices-my.sharepoint.com/:x:/r/personal/"
    "it-prod-bot2_kbs-services_com/_layouts/15/doc2.aspx"
    "?sourcedoc=%7B40BAF068-6659-43B8-B984-C1FA0C769CDD%7D"
    "&file=La%20Jolla%20Office%20Parking%20Validation.xlsx&action=edit"
    "&mobileredirect=true"
)


def retry_goto(page, url, max_retries=5):
    strategies = ["domcontentloaded", "commit", "load", "networkidle"]
    last_error = None
    for attempt in range(1, max_retries + 1):
        strategy = strategies[(attempt - 1) % len(strategies)]
        try:
            page.goto(url, wait_until=strategy, timeout=30000)
            page.wait_for_timeout(2000)
            return
        except Exception as e:
            last_error = e
            if attempt < max_retries:
                time.sleep(3 + random.uniform(0, 2))
    raise Exception(f"Navigation failed after {max_retries} attempts: {last_error}")


def authenticate(page, context, email, password):
    print("Authenticating with Microsoft...")
    for _ in range(3):
        email_el = page.locator("#i0116").first
        if email_el.is_visible(timeout=2000):
            email_el.fill(email)
            page.locator("#idSIButton9, input[type=submit]").first.click()
            page.wait_for_timeout(2000)
            continue

        password_el = page.locator("#i0118").first
        if password_el.is_visible(timeout=2000):
            password_el.fill(password)
            page.locator("#idSIButton9, input[type=submit]").first.click()
            page.wait_for_timeout(2000)
            continue

        stay = page.locator("#idSIButton9").first
        if stay.is_visible(timeout=2000):
            stay.click()
            page.wait_for_timeout(2000)
            continue
        break

    context.storage_state(path=str(AUTH_FILE))
    print(f"Session saved to {AUTH_FILE}")


def extract_download_url(html):
    pattern = re.compile(
        r'(download\.aspx\?UniqueId=[0-9a-f-]+'
        r'(?:\\u0026|&amp;)[^"\'<>\s]+tempauth[^"\'<>\s]+)'
    )
    match = pattern.search(html)
    if not match:
        return None
    base = "https://kbsservices-my.sharepoint.com/personal/it-prod-bot2_kbs-services_com/_layouts/15/"
    return base + match.group(1).replace("\\u0026", "&").replace("&amp;", "&")


def is_valid_xlsx(path):
    try:
        with zipfile.ZipFile(path) as z:
            return "xl/workbook.xml" in z.namelist()
    except Exception:
        return False


def append_validated(ticket_number, status, name, reason):
    now = datetime.now().isoformat(timespec="seconds")
    new_row = pd.DataFrame([{
        "Ticket Number": str(ticket_number),
        "Status": status,
        "Validated At": now,
        "Submitter Name": name or "anonymous",
        "Reason": reason or "",
    }])
    new_row.to_csv(VALIDATED_CSV, mode="a", header=not VALIDATED_CSV.exists(), index=False)
    print(f"  \u2713 {ticket_number} \u2192 {status}")


def send_email(ticket, status, name, reason, smtp_email, smtp_password):
    if not smtp_email or not smtp_password:
        print("  Email skipped (SMTP credentials not set)")
        return

    bcc = "Obed.Espinoza@kbs-services.com"
    subject = f"Parking Validation \u2014 Ticket {ticket} \u2014 {status}"
    body = (
        f"Ticket: {ticket}\n"
        f"Submitter: {name}\n"
        f"Status: {status}\n"
        f"Reason: {reason}\n"
        f"Time: {datetime.now().isoformat(timespec='seconds')}\n"
    )

    msg = MIMEText(body)
    msg["Subject"] = subject
    msg["From"] = smtp_email
    msg["To"] = bcc

    try:
        with smtplib.SMTP("smtp.office365.com", 587, timeout=30) as server:
            server.starttls()
            server.login(smtp_email, smtp_password)
            server.sendmail(smtp_email, [bcc], msg.as_string())
        print(f"  Email sent ({ticket} \u2192 {status})")
    except Exception as e:
        print(f"  Email failed ({ticket}): {e}", file=sys.stderr)


def send_sms(ticket, status, name, reason, phone, twilio_sid, twilio_token, twilio_msg_svc):
    if not twilio_sid or not twilio_token or not twilio_msg_svc:
        print("  SMS skipped (Twilio Messaging Service not set)")
        return
    if not phone:
        print(f"  SMS skipped (no phone number for ticket {ticket})")
        return
    try:
        num = f"+1{int(phone)}"
        body = f"Parking {ticket}: {status}. {name}, {reason}"
        Client(twilio_sid, twilio_token).messages.create(
            body=body, messaging_service_sid=twilio_msg_svc, to=num
        )
        print(f"  SMS sent ({ticket} \u2192 {num})")
    except Exception as e:
        print(f"  SMS failed ({ticket}): {e}", file=sys.stderr)


def validate_tickets(tickets, flash_email, flash_password, smtp_email, smtp_password, twilio_sid, twilio_token, twilio_msg_svc):
    flash_login_url = "https://v.flashvalet.com/secure/login.aspx"
    flash_main_url = "https://v.flashvalet.com/secure/Main.aspx"

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=True,
            channel="chrome",
            args=["--ignore-certificate-errors", "--disable-http2"],
        )
        page = browser.new_page()

        print("Logging into Flash Valet...")
        retry_goto(page, flash_login_url)

        page.locator("#ctl00_ContentPlaceHolder_UserNameTextBoxText").fill(flash_email)
        page.locator("#ctl00_ContentPlaceHolder_PasswordTextBoxText").fill(flash_password)
        page.locator("#ctl00_ContentPlaceHolder_btnLoginImageButton").click()
        page.wait_for_timeout(3000)

        if "login" in page.url.lower():
            print("Error: Flash Valet login failed", file=sys.stderr)
            sys.exit(1)

        print("Flash Valet login successful.")

        for t in tickets:
            ticket = str(t["Ticket Number (Bold Numbers in the Middle of Ticket)"])
            name = t.get("Name") or "anonymous"
            reason = t.get("Reason?") or ""
            phone = t.get("Phone number")

            print(f"Validating ticket {ticket}...")

            if "main" not in page.url.lower():
                page.goto(flash_main_url, wait_until="domcontentloaded", timeout=30000)
                page.wait_for_timeout(2000)

            ticket_input = page.locator("#ctl00_ContentPlaceHolder_TicketNumberInputBoxText")
            ticket_input.wait_for(timeout=10000)
            ticket_input.clear()
            ticket_input.fill(ticket)

            page.keyboard.press("Enter")
            page.wait_for_timeout(2000)

            dialog = page.locator("#dialog")
            if dialog.is_visible(timeout=3000):
                dialog_text = dialog.inner_text()
                if "No valid ticket found" in dialog_text:
                    print(f"  \u2717 Ticket {ticket} not found")
                    append_validated(ticket, "Invalid", name, reason)
                    send_email(ticket, "Invalid", '', reason, smtp_email, smtp_password)
                    send_sms(ticket, "Invalid", '', reason, phone, twilio_sid, twilio_token, twilio_msg_svc)
                    ok_btn = page.locator("button:has-text('OK'), a:has-text('OK')").first
                    if ok_btn.is_visible(timeout=2000):
                        ok_btn.click()
                    continue

            page.locator("#ctl00_ContentPlaceHolder_ValidationPriceInputBoxList").select_option("Full Validation")
            page.wait_for_timeout(500)

            page.locator("#ctl00_ContentPlaceHolder_SaveButtonImageButton").first.click()
            page.wait_for_timeout(2000)

            msg_dialog = page.locator("#dialog")
            if msg_dialog.is_visible(timeout=3000):
                print(f"  Message: {msg_dialog.inner_text().strip()}")

            ok_btn = page.locator("button:has-text('OK'), a:has-text('OK')").first
            if ok_btn.is_visible(timeout=5000):
                ok_btn.click()
                page.wait_for_timeout(1000)

            append_validated(ticket, "Validated", name, reason)
            send_email(ticket, "Validated", '', reason, smtp_email, smtp_password)
            send_sms(ticket, "Validated", '', reason, phone, twilio_sid, twilio_token, twilio_msg_svc)

        browser.close()


def main():
    try:
        from dotenv import load_dotenv
        load_dotenv()
    except Exception:
        pass

    def _cred(key, default=""):
        try:
            return get_secret(key) or os.environ.get(key, default)
        except Exception:
            return os.environ.get(key, default)

    EMAIL = _cred("ONELOGIN_EMAIL_it-prod-bot2")
    PASSWORD = _cred("ONELOGIN_PASSWORD_it-prod-bot2")
    FLASHVALET_EMAIL = _cred("FLASHVALET_EMAIL")
    FLASHVALET_PASSWORD = _cred("FLASHVALET_PASSWORD")
    SMTP_EMAIL = _cred("it-prod-bot-EMAIL")
    SMTP_PASSWORD = _cred("it-prod-bot-PASSWORD")
    TWILIO_SID = _cred("TWILIO_ACCOUNT_SID")
    TWILIO_TOKEN = _cred("TWILIO_AUTH_TOKEN")
    TWILIO_MSG_SVC = _cred("TWILIO_MESSAGING_SERVICE_SID")

    if not EMAIL or not PASSWORD:
        print(
            "Set ONELOGIN_EMAIL_it-prod-bot2 and ONELOGIN_PASSWORD_it-prod-bot2 in vault or .env",
            file=sys.stderr,
        )
        sys.exit(1)

    DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)

    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(
                headless=True,
                channel="chrome",
                args=["--ignore-certificate-errors", "--disable-http2"],
            )

            storage = str(AUTH_FILE) if AUTH_FILE.exists() else None
            context = browser.new_context(storage_state=storage)
            page = context.new_page()

            print("Opening SharePoint...")
            retry_goto(page, SHAREPOINT_URL)

            if "login" in page.url.lower():
                authenticate(page, context, EMAIL, PASSWORD)

            page.wait_for_timeout(4000)

            print("Extracting download URL...")
            for _ in range(5):
                html = page.content()
                dl_url = extract_download_url(html)
                if dl_url:
                    break
                time.sleep(2)

            if not dl_url:
                print(
                    f"Error: download URL not found. Page: {page.url[:80]}",
                    file=sys.stderr,
                )
                sys.exit(1)

            print("Downloading file...")
            page.evaluate("url => { location.href = url; }", dl_url)

            with page.expect_download(timeout=30000) as info:
                pass
            download = info.value
            file_path = DOWNLOAD_DIR / download.suggested_filename
            download.save_as(str(file_path))

            browser.close()

        if not is_valid_xlsx(file_path):
            print(
                "Error: downloaded file is not a valid Excel workbook",
                file=sys.stderr,
            )
            sys.exit(1)

        print(f"Downloaded: {file_path} ({file_path.stat().st_size} bytes)")

        print("Checking for new tickets...")

        validated = pd.read_csv(VALIDATED_CSV)
        validated_tickets = validated["Ticket Number"].astype(str).str.strip()

        xl = pd.read_excel(file_path, sheet_name="Sheet1")
        xl_tickets = xl["Ticket Number (Bold Numbers in the Middle of Ticket)"].astype(str).str.strip()

        new_df = xl[~xl_tickets.isin(validated_tickets)].copy()

        if new_df.empty:
            print("No new tickets to validate.")
        else:
            print(f"{len(new_df)} new ticket(s) found:")
            for _, row in new_df.iterrows():
                print(
                    f"  Ticket {row['Ticket Number (Bold Numbers in the Middle of Ticket)']}"
                    f" \u2014 {row['Name'] or 'anonymous'} ({row['Reason?'] or 'N/A'})"
                )

        new_tickets = new_df.to_dict("records")

        if new_tickets and FLASHVALET_EMAIL and FLASHVALET_PASSWORD:
            validate_tickets(new_tickets, FLASHVALET_EMAIL, FLASHVALET_PASSWORD, SMTP_EMAIL, SMTP_PASSWORD, TWILIO_SID, TWILIO_TOKEN, TWILIO_MSG_SVC)
        elif new_tickets and (not FLASHVALET_EMAIL or not FLASHVALET_PASSWORD):
            print("Warning: new tickets found but FLASHVALET credentials not set in vault or .env", file=sys.stderr)

    except Exception:
        traceback.print_exc(file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
