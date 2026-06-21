orders = [
    {"po": "001", "supplier": "RS", "cost": "25.50"},
    {"po": "002", "supplier": "IFM", "cost": "80.00"},
    {"po": "003", "supplier": "RS", "cost": "12.99"},
    {"po": "004", "supplier": "IFM", "cost": "20.00"}
]

supplier_totals = {}


for order in orders:
    supplier = order["supplier"]
    cost = float(order["cost"])
    if supplier not in supplier_totals:
        supplier_totals[supplier] = cost
    else:
        supplier_totals[supplier] += cost

#print(supplier_totals)

for supplier, total in supplier_totals.items():
    print(f"{supplier} total: £{total:.2f}")