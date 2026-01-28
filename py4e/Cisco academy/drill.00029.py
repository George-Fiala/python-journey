inventory = [
    {"name": "Screw M6", "price": 2.5, "qty": 100},
    {"name": "Nut M6", "price": 1.2, "qty": 500},
    {"name": "Washer", "price": 0.5, "qty": 1000},
    {"name": "Sensor", "price": 150.0, "qty": 5}
]

total_warehouse_value = 0

print("-" * 30)
print("INVENTORY REPORT")
print("-" * 30)

for item in inventory:
    row_value = item["price"] * item["qty"]

    total_warehouse_value += row_value

    print(f"{item["name"]}: {row_value}")

    print("-" * 30)
    print(f"Total value: {total_warehouse_value}")
   
    