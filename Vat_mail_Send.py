import pandas as pd
import smtplib
from email.message import EmailMessage
import os
from deep_translator import GoogleTranslator

# CONFIG
EMAIL_USER = "xyz@gmail.com"
EMAIL_PASS = "88888888"
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587

ORDERS_FOLDER = "customer_files_EU"

# Language Map
LANG_MAP = {
    "DE": "de", "AT": "de", "CH": "de", "LI": "de", "LU": "de",
    "FR": "fr", "BE": "fr", "MC": "fr",
    "ES": "es",
    "IT": "it", "SM": "it", "VA": "it",
    "NL": "nl",
    "PL": "pl",
    "PT": "pt",
    "SE": "sv",
    "DK": "da",
    "FI": "fi",
    "NO": "no",
    "CZ": "cs",
    "SK": "sk",
    "HU": "hu",
    "RO": "ro",
    "BG": "bg",
    "GR": "el",
    "EE": "et",
    "LV": "lv",
    "LT": "lt",
    "SI": "sl",
    "HR": "hr",
    "IE": "en",
    "GB": "en",
    "MT": "en",
    "IS": "is",
    "AL": "sq",
    "MK": "mk",
    "RS": "sr",
    "ME": "sr",
    "BA": "bs",
    "UA": "uk",
    "BY": "be",
    "MD": "ro",
    "DEFAULT": "en"
}

# -------------------------------
# ✅ English Template
# -------------------------------
def get_english_text(customer_id, vat_number, amount, store_name):
    return f"""
We hope this message finds you well.

We would like to inform you that, due to an internal system error, VAT was not charged at the time of the transaction, although it should have been charged.

Furthermore, upon verification through the VIES (European Commission) system, we noted that the VAT number provided ({vat_number}) appears to be invalid. Consequently, the VAT amount was not applied.

We wish to bring to your attention that an outstanding VAT amount of €{amount} remains to be collected.

Please find the relevant details attached for your review. We kindly request you to examine the information and arrange for the remittance of the applicable VAT amount at your earliest convenience.

Alternatively, if you believe that VAT should not be applicable, we kindly request you to provide a valid VAT number so that we may verify and resolve this matter accordingly.

If you are not the primary contact for this account, we would appreciate it if you could share the details of the appropriate contact person.

Company: {store_name}
Customer ID: {customer_id}
VAT number: {vat_number}

Should you require any further clarification, please do not hesitate to contact us.

Our bank details are as follows:
IBAN: NL79RABO0388440953
BIC: RABONL2U
KVK: 93127987.

Thank you for your attention to this matter and your cooperation.

Yours sincerely,  
Finance Team
"""

# -------------------------------
# 🌍 Translate Function
# -------------------------------
def translate_text(text, lang):
    try:
        if lang == "en":
            return ""  # skip translation
        return GoogleTranslator(source='en', target=lang).translate(text)
    except Exception as e:
        print(f"⚠️ Translation failed: {e}")
        return ""

# -------------------------------
# 📩 Final Email Body
# -------------------------------
def build_email_body(country_code, customer_id, vat, amount, store_name):
    lang = LANG_MAP.get(country_code, LANG_MAP["DEFAULT"])

    english_text = get_english_text(customer_id, vat, amount, store_name)
    local_text = translate_text(english_text, lang)

    if local_text:
        return f"""
{local_text}

------------------------------------------------------------

{english_text}
"""
    else:
        return english_text


# -------------------------------
# 📊 Load Excel
# -------------------------------
df = pd.read_excel("VAT Mail Data.xlsx")

# -------------------------------
# 📧 Send Emails
# -------------------------------
with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
    server.starttls()
    server.login(EMAIL_USER, EMAIL_PASS)

    for _, row in df.iterrows():
        try:
            customer_id = str(row["Customer ID"])
            email = row["Email Id"]
            country = row["Country"]
            vat = row["VAT number"]
            amount = round(float(row["VAT amount"]), 2)
            store_name = row.get("Company")

            msg = EmailMessage()
            msg["Subject"] = f"VAT Payment Request - {customer_id}"
            msg["From"] = EMAIL_USER
            msg["To"] = email
            msg["Cc"] = row.get("CC", "")

            # ✅ NEW BODY (Local + English)
            msg.set_content(
                build_email_body(country, customer_id, vat, amount, store_name)
            )

            # 📎 Attach CSV
            file_path = os.path.join(ORDERS_FOLDER, f"{customer_id}.csv")

            if os.path.exists(file_path):
                with open(file_path, "rb") as f:
                    msg.add_attachment(
                        f.read(),
                        maintype="application",
                        subtype="octet-stream",
                        filename=f"{customer_id}.csv"
                    )
            else:
                print(f"⚠️ File not found for {customer_id}")

            # 🚀 Send
            server.send_message(msg)
            print(f"✅ Sent: {email}")

        except Exception as e:
            print(f"❌ Error for {email}: {e}")