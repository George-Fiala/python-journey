supplier_codes = [
    {"supplier_c":"SUP01"},
    {"supplier_c":"SUP02"},
    {"supplier_c":"SUP03"},
    {"supplier_c":"SUP02"},
    {"supplier_c":"SUP04"},
    {"supplier_c":"SUP03"},
    {"supplier_c":"SUP03"}
]


seen = []
duplicates = []


for supplier_code in supplier_codes:
    supplier_c = supplier_code.get("supplier_c")
    if supplier_c in seen:
        if supplier_c not in duplicates:
            duplicates.append(supplier_c)
    else:
        seen.append(supplier_c)


print(f"Seen: {seen}")
print(f"Duplicates: {duplicates}")