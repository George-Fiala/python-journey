parts = [
    {"part_code": "BRG001", "description": "Bearing 6205", "quantity": 2, "unit_price": 25.50},
    {"part_code": "SNS002", "description": "IFM sensor", "quantity": 7, "unit_price": 80.00},
    {"part_code": "BLT003", "description": "Belt A32", "quantity": 0, "unit_price": 12.99},
    {"part_code": "MTR004", "description": "Motor 0.75kW", "quantity": 1, "unit_price": 120.00}
]
total_stock_value = 0
for part in parts:
    part_code = part["part_code"]
    quantity = part["quantity"]
    unit_price = part["unit_price"]
    item_value = quantity * unit_price
    
    print(f"{part_code} value: £{item_value:.2f}")
    total_stock_value += item_value
print(f"Total value: £{total_stock_value:.2f}")