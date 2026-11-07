deliveries = [
    {"supplier": "NORD", "delivery_days": 4, "received": True},
    {"supplier": "ERIKS", "delivery_days": 2, "received": True},
    {"supplier": "NORD", "delivery_days": 8, "received": True},
    {"supplier": "Mosca", "delivery_days": 3, "received": True},
    {"supplier": "ERIKS", "delivery_days": 6, "received": True},
    {"supplier": "Rubix", "delivery_days": 2, "received": False},
    {"supplier": "NORD", "delivery_days": 1, "received": False}
]

def build_delivery_summary(deliveries):
    summary = {}
    for delivery in deliveries:
        supplier = delivery["supplier"]
        delivery_days = delivery["delivery_days"]
        received = delivery["received"]
        if received:
            if supplier not in summary:
                summary[supplier] = {"count": 0, "total_days": 0}
            summary[supplier]["count"] += 1
            summary[supplier]["total_days"] += delivery_days
    return summary


delivery_summary = build_delivery_summary(deliveries)


if delivery_summary:
    for supplier in delivery_summary:
        details = delivery_summary[supplier]
        average = details["total_days"] / details["count"]
        print(f"{supplier} - Count: {details["count"]} Average: {average:.2f} days")

else:
    print("No deliveries")