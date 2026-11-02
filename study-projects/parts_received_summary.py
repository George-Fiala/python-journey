deliveries = [
    {"supplier": "NORD", "quantity": 5, "received": True},
    {"supplier": "ERIKS", "quantity": 3, "received": True},
    {"supplier": "NORD", "quantity": 8, "received": True},
    {"supplier": "Rubix", "quantity": 6, "received": False},
    {"supplier": "ERIKS", "quantity": 4, "received": True}
]

def build_received_summary(deliveries):
    summary ={}
    for delivery in deliveries:
        supplier = delivery["supplier"]
        quantity = delivery["quantity"]
        received = delivery["received"]
        if received:
            if supplier not in summary:
                summary[supplier] = {"count": 0, "total_quantity": 0}
            summary[supplier]["count"] += 1
            summary[supplier]["total_quantity"] += quantity
    return summary


received_summary = build_received_summary(deliveries)

if received_summary:
    for supplier in received_summary:
        details = received_summary[supplier]
        print(f"{supplier} - count: {details["count"]} - quantity: {details["total_quantity"]}")

else:
    print("No deliveries")