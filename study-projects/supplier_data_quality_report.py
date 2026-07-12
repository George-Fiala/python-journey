suppliers = [
    {"supplier_name": "EMBA", "supplier_category": "OEM Machinery", "account_number": None, "main_email": "service@emba-test.com", "main_phone": None, "is_24_7": False},
    {"supplier_name": "Pack Support", "supplier_category": "Packaging Machinery and Spare Parts", "account_number": "PS001", "main_email": "info@packsupport-test.com", "main_phone": "+44 0000 111111", "is_24_7": False},
    {"supplier_name": "D2", "supplier_category": "Bearings and Engineering", "account_number": None, "main_email": None, "main_phone": "+44 0000 222222", "is_24_7": True},
    {"supplier_name": "Techbelt", "supplier_category": "Conveyor Belting", "account_number": "TB001", "main_email": "sales@techbelt-test.com", "main_phone": "+44 0000 333333", "is_24_7": True},
    {"supplier_name": "Vulcatech", "supplier_category": "Conveyor Belting", "account_number": None, "main_email": None, "main_phone": None, "is_24_7": False}

]

total_suppliers = 0
missing_account_numbers = 0
missing_emails = 0
missing_phones = 0
suppliers_24_7 = 0


for supplier in suppliers:
    total_suppliers += 1
    account_number = supplier["account_number"]
    if account_number is None:
        missing_account_numbers += 1
    main_email = supplier["main_email"]
    if main_email is None:
        missing_emails += 1
    main_phone = supplier["main_phone"]
    if main_phone is None:
        missing_phones += 1
    is_24_7 = supplier["is_24_7"]
    if is_24_7 is True:
        suppliers_24_7 += 1
print("Supplier data quality report: ")
print(f"Total suppliers: {total_suppliers}")
print(f"Missing account numbers: {missing_account_numbers}")
print(f"Missing emails: {missing_emails}")
print(f"Missing phones: {missing_phones}")
print(f"Suppliers that are 24/7: {suppliers_24_7}")

