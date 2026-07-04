orders = [
    {"order_no": "PO001", "supplier": "RS", "cost": 120.50},
    {"order_no": "PO002", "supplier": "IFM", "cost": 480.00},
    {"order_no": "PO003", "supplier": "RS", "cost": 205.00},
    {"order_no": "PO004", "supplier": "Brammer", "cost": 120.00},
    {"order_no": "PO005", "supplier": "RS", "cost": 55.25}
]

supplier_totals = {}

for order in orders:
    supplier = order["supplier"]
    cost = order["cost"]
    if supplier not in supplier_totals:
        supplier_totals[supplier] = 0
    supplier_totals[supplier] += cost
for supplier, total in supplier_totals.items():    
    print(f"{supplier}: £{total:.2f}")