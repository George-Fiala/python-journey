orders = [
    {"po": "GF260401", "machine": "Emba 2", "value": 180, "delivered": False},
    {"po": "GF260402", "machine": "Bobst 1", "value": 540, "delivered": False},
    {"po": "GF260403", "machine": "Emba 2", "value": 120, "delivered": True},
    {"po": "GF260404", "machine": "Emba 2", "value": 320, "delivered": False},
    {"po": "GF260405", "machine": "Bobst 1", "value": 60, "delivered": False},
    {"po": "GF260406", "machine": "Gopfert 1", "value": 250, "delivered": True}
]


def build_pending_po_summary(orders):
    undelivered_orders = {}
    for order in orders:
        machine = order["machine"]
        value = order["value"]
        delivered = order["delivered"]
        if not delivered:
            if machine not in undelivered_orders:
                undelivered_orders[machine] = {"count": 0, "total_value": 0}
            undelivered_orders[machine]["count"] += 1
            undelivered_orders[machine]["total_value"] += value

    return undelivered_orders



pending_orders = build_pending_po_summary(orders)

if pending_orders:
    print(pending_orders)

else:
    print("No pending orders")
