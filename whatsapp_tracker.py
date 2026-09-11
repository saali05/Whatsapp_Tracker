import os
import re
import time
from datetime import datetime, timedelta
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By

# Known bank names and common keywords to exclude from Names
BANK_KEYWORDS = [
    "bank",
    "branch",
    "finance",
    "fedrel",
    "federal",
    "esaf",
    "sbi",
    "hdfc",
    "icici",
    "axis",
    "canara",
    "kerala",
    "gramin",
    "ബ്രാഞ്ച്",
    "account",
    "ifsc",
    "code",
]


def get_current_15min_filename():
    """Generates an Excel filename based on the current 15-minute slot."""
    now = datetime.now()

    # Round down minutes to the nearest 15-minute window (0, 15, 30, 45)
    start_minute = (now.minute // 15) * 15
    slot_start = now.replace(minute=start_minute, second=0, microsecond=0)
    slot_end = slot_start + timedelta(minutes=15)

    date_str = slot_start.strftime("%Y-%m-%d")
    start_str = slot_start.strftime("%H-%M")
    end_str = slot_end.strftime("%H-%M")

    return f"Account_Data_{date_str}_{start_str}_to_{end_str}.xlsx"


def extract_details(text):
    clean_text = text.replace("\r", "")
    lines = [line.strip() for line in clean_text.split("\n") if line.strip()]

    # 1. Extract IFSC Code
    ifsc_match = re.search(r"\b([A-Za-z]{4}0[A-Za-z0-9]{6})\b", clean_text)
    if not ifsc_match:
        ifsc_match = re.search(r"([A-Za-z]{4}0[A-Za-z0-9]{6})", clean_text)
    ifsc = ifsc_match.group(1).upper() if ifsc_match else ""

    # 2. Extract Account Number (9 to 18 consecutive digits)
    acc_no = ""
    labeled_acc = re.search(
        r"(?:a/c|acc(?:ount)?\.?|no\.?|number:?)\s*([0-9]{9,18})",
        clean_text,
        re.IGNORECASE,
    )
    if labeled_acc:
        acc_no = labeled_acc.group(1)
    else:
        digits = re.findall(r"\b([0-9]{9,18})\b", clean_text)
        if digits:
            acc_no = digits[0]

    # 3. Extract Amount
    amount = ""
    amount_match = re.search(
        r"(?:rs\.?|inr|amount:?)\s*([0-9]+(?:\.[0-9]{1,2})?)",
        clean_text,
        re.IGNORECASE,
    )
    if amount_match:
        amount = amount_match.group(1)
    else:
        for line in lines:
            if re.fullmatch(r"[0-9]{2,6}", line) and line != acc_no:
                amount = line
                break

    # 4. Extract Name
    name = ""
    labeled_name = re.search(
        r"(?:name|holder):?\s*([A-Za-z\s.]+?)(?=\s*(?:A/C|IFSC|acc|bank|\d|$))",
        clean_text,
        re.IGNORECASE,
    )
    if labeled_name:
        name = labeled_name.group(1).strip()
    else:
        inline_name = re.search(
            r"([A-Za-z\s.]{3,35})\s+A/C", clean_text, re.IGNORECASE
        )
        if inline_name:
            candidate = inline_name.group(1).strip()
            if not any(k in candidate.lower() for k in BANK_KEYWORDS):
                name = candidate

    if not name and len(lines) >= 2:
        for line in lines:
            line_lower = line.lower()
            if (
                re.search(r"\d", line)
                or (ifsc and ifsc.lower() in line_lower)
                or any(k in line_lower for k in BANK_KEYWORDS)
            ):
                continue
            if re.fullmatch(r"[A-Za-z\s.]{3,35}", line):
                name = line
                break

    if acc_no or ifsc:
        return {
            "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "Name": name,
            "Account Number": acc_no,
            "IFSC Code": ifsc,
            "Amount": amount,
            "Raw Text": clean_text,
        }
    return None


def save_to_excel(record):
    excel_file = get_current_15min_filename()
    columns = [
        "Timestamp",
        "Name",
        "Account Number",
        "IFSC Code",
        "Amount",
        "Raw Text",
    ]
    df_new = pd.DataFrame([record], columns=columns)
    df_new["Account Number"] = df_new["Account Number"].astype(str)

    if os.path.exists(excel_file):
        try:
            df_existing = pd.read_excel(
                excel_file, dtype={"Account Number": str}
            )
            df_combined = pd.concat([df_existing, df_new], ignore_index=True)
            df_combined.drop_duplicates(subset=["Raw Text"], inplace=True)
            df_combined.to_excel(excel_file, index=False)
        except PermissionError:
            print(
                f"[!] Error: Please close {excel_file} in Excel so it can be updated!"
            )
            return
    else:
        df_new.to_excel(excel_file, index=False)

    print(f"\n[+] SAVED TO: {excel_file}")
    print(f"    Name           : {record['Name']}")
    print(f"    Account Number : {record['Account Number']}")
    print(f"    IFSC Code      : {record['IFSC Code']}")
    print(f"    Amount         : {record['Amount']}\n")


def monitor_whatsapp():
    options = webdriver.ChromeOptions()
    profile_dir = os.path.abspath("whatsapp_session")
    options.add_argument(f"--user-data-dir={profile_dir}")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--remote-debugging-port=9222")
    options.add_argument("--disable-gpu")

    driver = webdriver.Chrome(options=options)
    driver.get("https://web.whatsapp.com")

    print("Opening WhatsApp Web...")
    time.sleep(15)
    print("Listening for messages...")

    seen_messages = set()

    while True:
        try:
            message_elements = driver.find_elements(
                By.CSS_SELECTOR,
                "div.message-in div.copyable-text, div.message-in span.selectable-text",
            )

            for el in message_elements:
                text = el.text.strip()
                if not text or len(text) < 6 or text in seen_messages:
                    continue

                parsed = extract_details(text)
                if parsed:
                    seen_messages.add(text)
                    save_to_excel(parsed)

        except Exception:
            pass

        time.sleep(2)


if __name__ == "__main__":
    monitor_whatsapp()