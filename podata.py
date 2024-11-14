import json
import pandas as pd

with open('hs 24_49_9H37_IMP_D889P0BII6_517365_2024-11_VEND_2024-05-31.pdf.json', 'r') as file:
    json_data = json.load(file)

fields = json_data.get("analyzeResult", {}).get("documents", [])[0].get("fields", {})
packed_item_table = fields.get("PACKING ITEM DETAILS", {}).get("valueArray", [])
distribution_centers_table = fields.get("DISTRIBUTION CENTRE DETAILS", {}).get("valueArray", [])
naming_series = "RTP-VPO-.YY.-.MM.-"
buyer = fields.get("BUYER", {}).get("valueString", "")
dept = fields.get("DEPT NO", {}).get("valueString", "")
import_po_number = fields.get("IMPORT PO NUMBER", {}).get("valueString", "")
pre_ticket = fields.get("PRE-TICKET", {}).get("valueString", "")
deal = fields.get("DEAL", {}).get("valueString", "")
freight_terms = fields.get("FREIGHT TERMS", {}).get("valueString", "")
vendor = fields.get("VENDOR", {}).get("valueString", "")
manufacturing_country = fields.get("MANUFACTURING COUNTRY", {}).get("valueString", "")
country_of_origin = fields.get("COUNTRY OF ORGIN", {}).get("valueString", "")
mode_of_transport = fields.get("MODE OF TRANSPORT", {}).get("valueString", "")
buyer_name = fields.get("BUYER NAME", {}).get("valueString", "")
store_no = fields.get("STORE NO", {}).get("valueString", "")
cost_currency = fields.get("COST CURRENCY", {}).get("valueString", "")
pre_ticket_instructions = fields.get("PRE-TICKET INSTRUCTIONS", {}).get("valueString", "")
exiting_country = fields.get("EXITING COUNTRY", {}).get("valueString", "")
exiting_port = fields.get("EXITING PORT", {}).get("valueString", "")
vendor_pack = fields.get("VENDOR PACK", {}).get("valueString", "")
total_units = fields.get("TOTAL UNITS", {}).get("valueString", "")
shipment_date = fields.get("CANCEL SHIPMENT DATE", {}).get("valueString", "")
start_ship_date = fields.get("START SHIP DATE", {}).get("valueString", "")
payment_days = fields.get("PAYMENT DAYS", {}).get("valueString", "")
payment_type = fields.get("PAYMENT TYPE", {}).get("valueString", "")
vendor_po = fields.get("PO", {}).get("valueString", "")
bill_address = fields.get("BILL ADDRESS", {}).get("valueString", "")
buyer_invoice_email = fields.get("BUYER INVOICE EMAIL", {}).get("valueString", "")
store_ready = fields.get("STORE READY", {}).get("valueString", "")

nest_code = [item.get("valueObject", {}).get("NEST CODE", {}).get("valueString", "") for item in packed_item_table]
styles = [item.get("valueObject", {}).get("VENDOR STYLE", {}).get("valueString", "") for item in packed_item_table]
types = [item.get("valueObject", {}).get("TYPE", {}).get("valueString", "") for item in packed_item_table]
unit_costs = [item.get("valueObject", {}).get("UNIT COST", {}).get("valueString", "") for item in packed_item_table]
detailed_descriptions = [item.get("valueObject", {}).get("DETAILED DESCRIPTION", {}).get("valueString", "") for item in packed_item_table]
quantities = [item.get("valueObject", {}).get("QUANTITY", {}).get("valueString", "") for item in packed_item_table]
sr_packs = [item.get("valueObject", {}).get("SR PACK", {}).get("valueString", "") for item in packed_item_table]
packs = [item.get("valueObject", {}).get("PACKS", {}).get("valueString", "") for item in packed_item_table]

store_numbers = [dc.get("valueObject", {}).get("STORE NO", {}).get("valueString", "") for dc in distribution_centers_table]
store_addresses = [dc.get("valueObject", {}).get("STORE ADDRESS", {}).get("valueString", "") for dc in distribution_centers_table]
cities = [dc.get("valueObject", {}).get("CITY", {}).get("valueString", "") for dc in distribution_centers_table]
provinces = [dc.get("valueObject", {}).get("PROVINCE", {}).get("valueString", "") for dc in distribution_centers_table]
postalcodes = [dc.get("valueObject", {}).get("POSTAL CODE", {}).get("valueString", "") for dc in distribution_centers_table]

df_items = pd.DataFrame({
    "Nest Code": nest_code,
    "Style": styles,
    "Type": types,
    "Unit Cost": unit_costs,
    "Detailed Description": detailed_descriptions,
    "Quantity": quantities,
    "SR Pack": sr_packs,
    "Packs": packs
})

df_items['BUYER'] = buyer
df_items['DEPT NO'] = dept
df_items['DEAL'] = deal
df_items['FREIGHT TERMS'] = freight_terms
df_items['VENDOR'] = vendor
df_items['COUNTRY OF ORGIN'] = country_of_origin
df_items['MODE OF TRANSPORT'] = mode_of_transport
df_items['BUYER NAME'] = buyer_name
df_items['COST CURRENCY'] = cost_currency
df_items['PRE-TICKET'] = pre_ticket
df_items['VENDOR PACK'] = vendor_pack
df_items['STORE READY'] = store_ready
df_items['TOTAL UNITS'] = total_units
df_items['START SHIP DATE'] = start_ship_date
df_items['CANCEL SHIPMENT DATE'] = shipment_date
df_items['PAYMENT DAYS'] = payment_days
df_items['PAYMENT TYPE'] = payment_type

df_dc = pd.DataFrame({
    "Store No": store_numbers,
    "Store Address": store_addresses,
    "City": cities,
    "Province": provinces,
    "Postal Code": postalcodes
})

with pd.ExcelWriter("Hs5_data.xlsx", engine='xlsxwriter') as writer:
    df_items.to_excel(writer, sheet_name="Packed Items", index=False)
    df_dc.to_excel(writer, sheet_name="Distribution Centers", index=False)

print("Data for buyer {} has been saved to Hs5_data.xlsx".format(buyer))
