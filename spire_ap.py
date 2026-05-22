import os
import re
from dotenv import load_dotenv
from requests.auth import HTTPBasicAuth
import requests
import invoices as invoice_processor

# ENVIRONMENT VARIABLES
load_dotenv()
ROOT_URL = os.getenv("SPIRE_ROOT")
SPIRE_USER = os.getenv("SPIRE_UN")
SPIRE_PASS = os.getenv("SPIRE_PW")
SPIRE_LIMIT = 5000
print(f"Current User: {SPIRE_USER}")

auth = HTTPBasicAuth(SPIRE_USER, SPIRE_PASS)
headers = {"accept": "application/json"}

def get_purchase_orders():
    response = requests.get(f"{ROOT_URL}/purchasing/orders?sort=-modified&limit={SPIRE_LIMIT}", headers=headers, auth=auth)
    response.raise_for_status()
    return response.json()

def get_purchase_history():
    response = requests.get(f"{ROOT_URL}/purchasing/history?sort=-modified&limit={SPIRE_LIMIT}", headers=headers, auth=auth)
    response.raise_for_status()
    return response.json()

def get_ap_transactions():
    response = requests.get(f"{ROOT_URL}/ap/transactions?sort=-modified&limit={SPIRE_LIMIT}", headers=headers, auth=auth)
    response.raise_for_status()
    return response.json()

# Returns 
def compare_candidates_to_purchases():
    # Get a list of recently received invoices in inbox (past 24 hours)
    invoices = invoice_processor.process_invoices()
    if not invoices:
        print("No invoices found!")
        return

    # Get a list of purchase orders from Spire to compare to
    recent_purchases_orders = get_purchase_orders()["records"]
    recent_purchases_orders.extend(get_purchase_history()["records"])
    po_list = {po["number"].lstrip("0"): po for po in recent_purchases_orders}

    invoice_to_po_map = {}
    
    # If one of the PO candidates from the invoices is found, add the file path and PO info in a dict
    for invoice in invoices:
        for candidate in invoice["po_candidates"]:
            if candidate in po_list:
                invoice_to_po_map[invoice["file_path"]] = po_list[candidate]
                break
                
    return invoice_to_po_map

if __name__ == "__main__":
    print(compare_candidates_to_purchases())