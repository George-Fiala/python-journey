orders = [
    {"order_no": "PO001", "supplier": "RS", "cost": 120.50},
    {"order_no": "PO002", "supplier": "IFM", "cost": 480.00},
    {"order_no": "PO003", "supplier": "RS", "cost": 205.00},
    {"order_no": "PO004", "supplier": "Brammer", "cost": 120.00},
    {"order_no": "PO005", "supplier": "RS", "cost": 55.25}
]

supplier_counts = {}

for order in orders:
    supplier = order["supplier"]
    if supplier not in supplier_counts:
        supplier_counts[supplier] = 0
    supplier_counts[supplier] += 1
print("Total count of orders per supplier: ")

for supplier, count in supplier_counts.items():
    print(f"{supplier}: {count} orders")