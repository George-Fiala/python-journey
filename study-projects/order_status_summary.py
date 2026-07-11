orders = [
    {"order_no": "PO001", "supplier": "RS", "cost": 120.50, "status": "Pending"},    
    {"order_no": "PO002", "supplier": "IFM", "cost": 480.00, "status": "Delivered"},
    {"order_no": "PO003", "supplier": "RS", "cost": 205.00, "status": "Overdue"},
    {"order_no": "PO004", "supplier": "Brammer", "cost": 200.00, "status": "Pending"},
    {"order_no": "PO005", "supplier": "RS", "cost": 200.00, "status": "Delivered"}
]

status_counts = {}
status_totals = {}

for order in orders:
    status = order["status"]
    cost = order["cost"]
    if status not in status_counts:
        status_counts[status] = 0
        status_totals[status] = 0
    status_counts[status] += 1
    status_totals[status] += cost
print("Order count by status: ")
for status, count in status_counts.items():
    print(f"{status}: {count} orders")
print("Spend by status: ")
for status, total in status_totals.items():
    print(f"{status}: £{total:.2f}")