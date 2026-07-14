suppliers = [
    {"supplier_name": "EMBA", "supplier_category": "OEM Machinery", "account_number": None, "main_email": "service@emba-test.com", "main_phone": None, "is_24_7": False},
    {"supplier_name": "Pack Support", "supplier_category": "Packaging Machinery and Spare Parts", "account_number": "PS001", "main_email": "info@packsupport-test.com", "main_phone": "+44 0000 111111", "is_24_7": False},
    {"supplier_name": "D2", "supplier_category": "Bearings and Engineering", "account_number": None, "main_email": None, "main_phone": "+44 0000 222222", "is_24_7": True},
    {"supplier_name": "Techbelt", "supplier_category": "Conveyor Belting", "account_number": "TB001", "main_email": "sales@techbelt-test.com", "main_phone": "+44 0000 333333", "is_24_7": True},
    {"supplier_name": "Vulcatech", "supplier_category": "Conveyor Belting", "account_number": None, "main_email": None, "main_phone": None, "is_24_7": False}

]

suppliers_without_complete_contacts = 0


for supplier in suppliers:
    main_email = supplier["main_email"]
    main_phone = supplier["main_phone"]
    if main_email is None or main_phone is None:
        suppliers_without_complete_contacts += 1
        supplier_name = supplier["supplier_name"]
        print(supplier_name)
print(f"Suppliers without complete contacts: {suppliers_without_complete_contacts}")
       
