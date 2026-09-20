total_quantity = 0

reorder_list = [
    {"name": "Sensor", "quantity": 2},
    {"name": "Belt", "quantity": 3}
]

for item in reorder_list:
    name = item["name"]
    quantity = item["quantity"]
    total_quantity += quantity

number_of_parts = len(reorder_list)
print(f"Number of parts to order: {number_of_parts}")
print(f"Total quantity to order: {total_quantity}")