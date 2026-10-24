orders = [
    {"po": "GF260201", "machine": "Emba 2", "delivered": False},
    {"po": "GF260202", "machine": "Bobst 1", "delivered": False},
    {"po": "GF260203", "machine": "Emba 2", "delivered": True},
    {"po": "GF260204", "machine": "Emba 2", "delivered": False},
    {"po": "GF260205", "machine": "Bobst 1", "delivered": False},
    {"po": "GF260206", "machine": "Gopfert 1", "delivered": True}
]


def count_pending_by_machine(orders):
    counts = {}
    for order in orders:
        machine = order["machine"]
        delivered = order["delivered"]
        if not delivered:
            counts[machine] = counts.get(machine, 0) + 1
    return counts



pending_orders = count_pending_by_machine(orders)

print(pending_orders)
