orders = [
    {"po": "GF260301", "machine": "Emba 2", "value": 180, "delivered": False},
    {"po": "GF260302", "machine": "Bobst 1", "value": 540, "delivered": False},
    {"po": "GF260303", "machine": "Emba 2", "value": 120, "delivered": True},
    {"po": "GF260304", "machine": "Emba 2", "value": 320, "delivered": False},
    {"po": "GF260305", "machine": "Bobst 1", "value": 60, "delivered": False},
    {"po": "GF260306", "machine": "Gopfert 1", "value": 250, "delivered": True}

]


def calculate_pending_value(orders):
    undelivered_orders = {}
    for order in orders:
        machine = order["machine"]
        value = order["value"]
        delivered = order["delivered"]
        if not delivered:
            undelivered_orders[machine] = undelivered_orders.get(machine, 0) + value
    return undelivered_orders


pending_value = calculate_pending_value(orders)

if pending_value:
    print(pending_value)

else:
    print("No pending orders")