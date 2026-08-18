parts = [
    {"part_name": "Bearing", "stock": 2, "minimum": 5, "unit_price": 80, "critical": True},
    {"part_name": "Belt", "stock": 10, "minimum": 4, "unit_price": 25, "critical": False},
    {"part_name": "Sensor", "stock": 1, "minimum": 3, "unit_price": 120, "critical": True},
    {"part_name": "Motor", "stock": 4, "minimum": 4, "unit_price": 300, "critical": False}
]


def get_stock_value(part):
    stock = part["stock"]
    unit_price = part["unit_price"]
    return stock * unit_price


def get_part_status(part):
    minimum = part["minimum"]
    critical = part["critical"]
    stock = part["stock"]
    if stock < minimum and critical:
        return "Urgent reorder"
    elif stock < minimum:
        return "Reorder"
    return "Stock OK"


total_stock_value = 0
urgent_reorder_count = 0
stock_ok_count = 0

for part in parts:
    stock_value = get_stock_value(part)
    part_status = get_part_status(part)
    part_name = part["part_name"]
    total_stock_value += stock_value
    if part_status == "Urgent reorder":
        urgent_reorder_count += 1
    if part_status == "Stock OK":
        stock_ok_count += 1
    print(f"{part_name} - £{stock_value:.2f} - {part_status}")

urgent_reorder_percentage = urgent_reorder_count / len(parts) * 100 
average_stock_value = total_stock_value / len(parts)


print(f"Total stock value: £{total_stock_value:.2f}")
print(f"Urgent reorders: {urgent_reorder_count}")
print(f"Urgent reorder percentage: {urgent_reorder_percentage}")
print(f"Average stock value: £{average_stock_value:.2f}")
print(f"Stock OK parts: {stock_ok_count}")