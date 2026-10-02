orders = [
    {"order_po": "PO1001", "supplier": "NORD", "value": 850, "approved": True, "received": True},
    {"order_po": "PO1002", "supplier": "ERIKS", "value": 1200, "approved": False, "received": False},
    {"order_po": "PO1003", "supplier": "Rubix", "value": 450, "approved": True, "received": False},
    {"order_po": "PO1004", "supplier": "NORD", "value": 2000, "approved": True, "received": False}
]


def get_po_status(value, approved, received):
    if not approved:
        return "Approval needed"
    elif value >= 1500 and not received:
        return "High value pending"
    elif not received:
        return "Pending receipt"
    else:
        return "Complete"


approval_needed_count = 0
pending_receipt_count = 0
complete_count = 0
high_value_pending_count = 0


for order in orders:
    order_po = order["order_po"]
    supplier = order["supplier"]
    value = order["value"]
    approved = order["approved"]
    received = order["received"]
    status = get_po_status(value, approved, received)

    if status == "Approval needed":
        approval_needed_count += 1
    if status == "Pending receipt":
        pending_receipt_count += 1
    if status == "Complete":
        complete_count += 1
    if status == "High value pending":
        high_value_pending_count += 1
    

    print(f"{order_po} - {supplier} - {value} - {status}")
print(f"Approval needed orders count: {approval_needed_count}")
print(f"Orders pending receipt: {pending_receipt_count}")
print(f"Complete orders: {complete_count}")
print(f"High value pending order count: {high_value_pending_count}")