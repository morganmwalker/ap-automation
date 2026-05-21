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
print(f"Current User: {SPIRE_USER}")

auth = HTTPBasicAuth(SPIRE_USER, SPIRE_PASS)
headers = {"accept": "application/json"}

def get_purchase_orders():
    response = requests.get(f"{ROOT_URL}/purchasing/orders?sort=-modified&limit=1000", headers=headers, auth=auth)
    response.raise_for_status()
    return response.json()

def get_purchase_history():
    response = requests.get(f"{ROOT_URL}/purchasing/history?sort=-modified&limit=1000", headers=headers, auth=auth)
    response.raise_for_status()
    return response.json()

def get_ap_transactions():
    response = requests.get(f"{ROOT_URL}/ap/transactions?sort=-modified&limit=1000", headers=headers, auth=auth)
    response.raise_for_status()
    return response.json()

def compare_candidates_to_purchases():
    invoices = invoice_processor.process_invoices()
    recent_purchases_orders = get_purchase_orders()["records"]
    recent_purchases_orders.extend(get_purchase_history()["records"])
    po_list = [po["number"].lstrip("0") for po in recent_purchases_orders]
    print(po_list)

if __name__ == "__main__":
    compare_candidates_to_purchases()