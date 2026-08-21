orders = [
    {"po_number": "PO001", "quantity": 3, "unit_price": 80, "urgent": True},
    {"po_number": "PO002", "quantity": 5, "unit_price": 20, "urgent": False},
    {"po_number": "PO003", "quantity": 2, "unit_price": 300, "urgent": True}
]


def get_overall_order_price(order):
    quantity = order["quantity"]
    unit_price = order["unit_price"]
    return quantity * unit_price


def get_part_status(overall_order_price, urgent):
    if overall_order_price >= 500 and urgent:
        return "Priority"
    return "Normal"

total_value_of_all_orders = 0
priority_count = 0


for order in orders:
    overall_order_price = get_overall_order_price(order)
    urgent = order["urgent"]
    part_status = get_part_status(overall_order_price, urgent)
    po_number = order["po_number"]
    if part_status =="Priority":
        priority_count += 1
    total_value_of_all_orders += overall_order_price
    print(f"{po_number} - £{overall_order_price:.2f} - {part_status}")


average_order_value = total_value_of_all_orders / len(orders)
percentage_of_priority_orders = priority_count / len(orders) * 100

print(f"Total value of orders: £{total_value_of_all_orders:.2f}")
print(f"Average order value: £{average_order_value:.2f}")
print(f"Priority orders count: {priority_count}")
print(f"Percentage of priority orders: {percentage_of_priority_orders:.2f}%")