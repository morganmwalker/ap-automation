import os
import re
from dotenv import load_dotenv
from requests.auth import HTTPBasicAuth
import requests

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

if __name__ == "__main__":
    active_pos = get_purchase_orders()
    for po in active_pos["records"]:
        print(po["number"].lstrip("0"))
