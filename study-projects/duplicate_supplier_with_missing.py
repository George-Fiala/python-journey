suppliers =[
    {"supplier_code": "SUP01"},
    {"supplier_code": "SUP02"},
    {"supplier_code": None},
    {"supplier_code": "SUP03"},
    {"supplier_code": "SUP02"},
    {},
    {"supplier_code": "SUP04"},
    {"supplier_code": "SUP03"},
    {"supplier_code": "SUP03"},
    {"supplier_code": None}
]



seen = []
duplicates =[]
missing_count = 0


for supplier in suppliers:
    supplier_code = supplier.get("supplier_code")
    if supplier_code is not None:
        if supplier_code in seen:
            if supplier_code not in duplicates:
                duplicates.append(supplier_code)
        else:
            seen.append(supplier_code)
    else:
        missing_count += 1

print(f"Seen: {seen}")
print(f"Duplicates: {duplicates}")
print(f"MIssing count: {missing_count}")