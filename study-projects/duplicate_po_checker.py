po_numbers = [
    {"po_number": "PO1001"},
    {"po_number": "PO1002"},
    {"po_number": "PO1003"},
    {"po_number": "PO1002"},
    {"po_number": "PO1004"},
    {"po_number": "PO1001"},
    {"po_number": "PO1002"}
]


seen = []
duplicates = []


for po_num in po_numbers:
    po_number = po_num.get("po_number")
    if po_number in seen:
        if po_number not in duplicates:
            duplicates.append(po_number)

    else:
        seen.append(po_number)


print(f"Seen: {seen}")
print(f"Duplicates: {duplicates}")

