from datetime import datetime, timedelta
import os
import re
from dotenv import load_dotenv
from pypdf import PdfReader
import win32com.client

# ENVIRONMENT VARIABLES
load_dotenv()
account_name = os.getenv("ACCOUNTING_EMAIL")
print(f"Targeting Account: {account_name}")

# GLOBAL VARIABLES
# Our Spire POs are 5 digits
PO_LENGTH = 5

TEMP_DIR = os.path.join(
    os.environ["USERPROFILE"],
    "AppData",
    "Local",
    "Temp",
    "InvoiceProcessor",
)
os.makedirs(TEMP_DIR, exist_ok=True)

# Extracts text from a digital PDF file
def extract_text_from_pdf(pdf_path):
    try:
        reader = PdfReader(pdf_path)
        text = ""
        for page in reader.pages:
            page_text = page.extract_text(layout_mode="layout")
            if page_text:
                text += page_text + "\n"
        return text
    except Exception as e:
        print(f"Error reading PDF {pdf_path}: {e}")
        return ""

# Finds all unique, realistic PO numbers on the invoice
def find_all_po_candidates(text):
    if not text:
        return []
    #print(f"-------\n\n{text}")
    candidates = set()

    # Check first for a labeled PO number
    prefix_pattern = r"(?:po|p/o|purchase\s*order|order|ref)\s*[:#-]?\s*(\d{1,10}(?:-[A-Z]{2,3})?)"
    for match in re.finditer(prefix_pattern, text, re.IGNORECASE):
        cleaned = clean_po_candidate(match.group(1))
        if len(cleaned) == PO_LENGTH:
            candidates.add(cleaned)
    
    # If the first sweep failed, find consecutive digits of length PO_LENGTH
    if not candidates:
        broad_pattern = rf"\b\d{{{PO_LENGTH},{PO_LENGTH+1}}}(?:-[A-Z]{{2,3}})?\b"
        for match in re.finditer(broad_pattern, text):
            cleaned = clean_po_candidate(match.group(0))
            if len(cleaned) == PO_LENGTH:
                candidates.add(cleaned)

        # If clean matches found nothing, search in broken/merged text columns
        merged_pattern = rf"(\d{{{PO_LENGTH}}})(?:-[A-Z]{{2,3}})?\b"
        for match in re.finditer(merged_pattern, text):
            cleaned = clean_po_candidate(match.group(1))  # Strips out just the target digits from the blob
            if len(cleaned) == PO_LENGTH:
                candidates.add(cleaned)

    return candidates

def clean_po_candidate(raw_token):
    """Cleans a potential PO string by stripping initials and leading zeros."""
    # Split trailing initials if appended with a dash (e.g., '68496-PJ' -> '68496')
    base_string = raw_token.split("-")[0]
    # Remove leading zeros
    return base_string.lstrip("0")

def process_invoices():
    # Connect to local Outlook application
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")

    # Access the accounting inbox
    target_folder = namespace.Folders[account_name].Folders["Inbox"]
    messages = target_folder.Items

    # Sort messages by received time to ensure accurate filtering
    messages.Sort("[ReceivedTime]", True)

    # Filter for emails received in the last 24 hours
    one_day_ago = datetime.now() - timedelta(days=1)
    filter_string = (
        f"[ReceivedTime] >= '{one_day_ago.strftime('%m/%d/%Y %H:%M %p')}'"
    )
    recent_messages = messages.Restrict(filter_string)

    print(
        f"Scanning emails received since: {one_day_ago.strftime('%Y-%m-%d %H:%M')}\n"
    )
    print(f"{'Sender':<30} | {'Subject':<30} | {'PO Found':<15}")
    print("-" * 80)

    for message in recent_messages:
        if message.Attachments.Count == 0:
            continue

        for attachment in message.Attachments:
            if attachment.FileName.lower().endswith(".pdf"):
                temp_pdf_path = os.path.join(TEMP_DIR, attachment.FileName)
                attachment.SaveAsFile(temp_pdf_path)

                pdf_text = extract_text_from_pdf(temp_pdf_path)

                if os.path.exists(temp_pdf_path):
                    os.remove(temp_pdf_path)

                # Decision logic: Is it an invoice?
                is_invoice = re.search(
                    r"invoice|tax invoice|inv-", pdf_text, re.IGNORECASE
                )

                if is_invoice:
                    candidates = list(find_all_po_candidates(pdf_text))
                    #print(candidates)
                    if not candidates:
                        po_number = "Not Found"
                    elif len(candidates) == 1:
                        po_number = candidates[0]  # Perfect single match
                    else:
                        # Multiple matches found
                        po_number = f"Verify: {candidates}"

                    sender = str(message.SenderName)[:30]
                    subject = str(message.Subject)[:30]
                    print(
                        f"{sender:<30} | {subject:<30} | {po_number:<15}"
                    )

if __name__ == "__main__":
    process_invoices()
