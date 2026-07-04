parts = [
    {"part_code": "BRG001", "supplier": "RS", "description": "Bearing 6205", "current_stock": 2, "minimum_stock": 5},
    {"part_code": "SNS002", "supplier": "IFM", "description": "IFM Sensor", "current_stock": 7, "minimum_stock": 4},
    {"part_code": "BLT003", "supplier": "RS", "description": "Belt A32", "current_stock": 0, "minimum_stock": 3},
    {"part_code": "MTR004", "supplier": "Brammer", "description": "Motor 0.75kW", "current_stock": 1, "minimum_stock": 1},
    {"part_code": "FLT005", "supplier": "RS", "description": "Filter", "current_stock": 10, "minimum_stock": 6}
]

search_supplier = input("Enter supplier: ")
found = False

for part in parts:
    supplier = part["supplier"]
    current_stock = part["current_stock"]
    minimum_stock = part["minimum_stock"]
    part_code = part["part_code"]
    description = part["description"]
    if supplier == search_supplier and current_stock < minimum_stock:
        found = True
        print(f"{part_code} - {description} | Current: {current_stock} | Minimum: {minimum_stock}")

if found is False:
    print("No low stock parts found for this supplier.")