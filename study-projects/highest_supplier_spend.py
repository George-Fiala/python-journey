orders = [
    {"po": "GF260501", "supplier": "NORD", "value": 180, "delivered": True},
    {"po": "GF260502", "supplier": "ERIKS", "value": 450, "delivered": True},
    {"po": "GF260503", "supplier": "NORD", "value": 620, "delivered": True},
    {"po": "GF260504", "supplier": "ERIKS", "value": 300, "delivered": False},
    {"po": "GF260505", "supplier": "NORD", "value": 240, "delivered": True},
    {"po": "GF260506", "supplier": "Mosca", "value": 750, "delivered": True}
]



def find_highest_po_by_supplier(orders):
    highest_order = {}
    for order in orders:
        po = order["po"]
        supplier = order["supplier"]
        value = order["value"]
        delivered = order["delivered"]
        if delivered:
            if supplier not in highest_order:
                highest_order[supplier] = order
            elif value > highest_order[supplier]["value"]:
                highest_order[supplier] = order
    return highest_order
            


highest_po = find_highest_po_by_supplier(orders)


if highest_po:
    for supplier in highest_po:
        details = highest_po[supplier]
        print(f"{supplier} - PO: {details["po"]}, Supplier: {details["supplier"]}, Value: £{details["value"]:.2f}, Delivered: {details["delivered"]}")

else:
    print("No orders")