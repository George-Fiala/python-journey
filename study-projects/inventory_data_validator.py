parts = [
    {"name": "Bearing", "supplier": "NORD", "bay": "A12"},
    {"name": "Sensor", "supplier": None, "bay": "B04"},
    {"name": "Belt", "supplier": "ERIKS", "bay": None},
    {"name": None, "supplier": "Rubix", "bay": "C02"},
    {"name": "Chain", "supplier": None, "bay": None},
    {"name": "Gearbox", "bay": "C02"}
]


def get_record_status(name, supplier, bay):
    if name is None:
        return "Missing name"
    elif supplier is None and bay is None:
        return "Missing supplier and bay"
    elif supplier is None:
        return "Missing supplier"
    elif bay is None:
        return "Missing bay"
    else:
        return "Valid"



valid_count = 0
invalid_count = 0


for part in parts:
    name = part.get("name")
    supplier = part.get("supplier")
    bay = part.get("bay")
    record_status = get_record_status(name, supplier, bay)
    if record_status == "Valid":
        valid_count += 1
    else:
        invalid_count += 1

    print(f"{name} - {supplier} - {bay} - {record_status}")
print(f"Valid parts: {valid_count}")
print(f"Invalid parts: {invalid_count}")