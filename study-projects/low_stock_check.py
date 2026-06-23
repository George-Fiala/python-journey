parts = [
    {"part_number": "BRG001", "description": "Bearing 6205", "current_stock": 2, "minimum_stock": 5},
    {"part_number": "SNS002", "description": "IFM Sensor", "current_stock": 7, "minimum_stock": 4},
    {"part_number": "BLT003", "description": "Belt A32", "current_stock": 0, "minimum_stock": 3},
    {"part_number": "MTR004", "description": "Motor 0.75kW", "current_stock": 1, "minimum_stock": 1}
]

for part in parts:
    part_number = part["part_number"]
    current_stock = part["current_stock"]
    minimum_stock = part["minimum_stock"]

    if current_stock < minimum_stock:
        print(f"{part_number} needs reorder. Current: {current_stock} | Minimum: {minimum_stock}")
    else:
        print(f"{part_number} is ok. Current: {current_stock} | Minimum: {minimum_stock}")