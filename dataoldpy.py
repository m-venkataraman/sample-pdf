import json
import pandas as pd

with open('TJX UK 55_61_5G994_188953_D653P06HRT_VEND_2024-12_10-07-2024_.pdf.json', 'r') as file:
    json_data = json.load(file)

fields = json_data.get("analyzeResult", {}).get("documents", [])[0].get("fields", {})
packed_item_table = json_data["analyzeResult"]["documents"][0]["fields"]["Packed Item Table"]["valueArray"]

buyer = fields.get("Buyer", {}).get("valueString")
dept = fields.get("DEPT", {}).get("valueString")
po = fields.get("PO", {}).get("valueString")
pre_ticket = fields.get("PRE-TICKET", {}).get("valueString")
store_ready_sr = fields.get("STORE READY (SR)", {}).get("valueString")
vendor_styles = [item["valueObject"]["VENDOR STYLE"]["valueString"] for item in packed_item_table]
item_descriptions = [item["valueObject"]["DESCRIPTION"]["valueString"] for item in packed_item_table]
ens_coord = [item["valueObject"]["ENS/COORD"]["valueString"] for item in packed_item_table]

df = pd.DataFrame({
    "Vendor Style": vendor_styles,
    "Description": item_descriptions,
    "Nest Code": ens_coord
})

df['Buyer'] = buyer
df['PRE-TICKET'] = pre_ticket
df['PO'] = po
df['DEPT'] = dept

df.to_excel("TJX-UK_data.xlsx", index=False)

print(buyer)
