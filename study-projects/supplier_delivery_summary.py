deliveries = [
    {"delivery": "D101", "supplier": "NORD", "value": 120, "received": True},
    {"delivery": "D102", "supplier": "ERIKS", "value": 300, "received": True},
    {"delivery": "D103", "supplier": "NORD", "value": 80, "received": False},
    {"delivery": "D104", "supplier": "NORD", "value": 250, "received": True},
    {"delivery": "D105", "supplier": "ERIKS", "value": 60, "received": True},
    {"delivery": "D106", "supplier": "Rubix", "value": 500, "received": False},
    {"delivery": "D107", "supplier": "Mosca", "value": 90, "received": True},
]


def build_delivery_summary(deliveries):
    inner_delivery_summary = {}
    for delivery in deliveries:
        supplier = delivery["supplier"]
        value = delivery["value"]
        received = delivery["received"]
        if received:
            if supplier not in inner_delivery_summary:
                inner_delivery_summary[supplier] = {"count": 0, "total_value": 0}
            inner_delivery_summary[supplier]["count"] += 1
            inner_delivery_summary[supplier]["total_value"] += value
    return inner_delivery_summary


delivery_summary = build_delivery_summary(deliveries)

if delivery_summary:
    for supplier in delivery_summary:
        details = delivery_summary[supplier]
        print(f"{supplier}: {details["count"]} deliveries -  £{details["total_value"]:.2f}")

else:
    print("No deliveries")