allowed_cost_centers = ["321A", "321B", "321C", "334"]


parts = [
    {"name": "Motor", "cost_centre": "321A"},
    {"name": "Sensor", "cost_centre": "334"},
    {"name": "Belt", "cost_centre": "999X"},
    {"name": "Bearing", "cost_centre": None},
    {"name": "Cylinder", "cost_centre": "321C"}
]


def get_cost_centre_status(cost_centre):
    if cost_centre is None:
        return "Missing"
    elif cost_centre in allowed_cost_centers:
        return "Valid"
    else: 
        return "Invalid"


valid_count = 0
invalid_count = 0
missing_count = 0


for part in parts:
    name = part.get("name")
    cost_centre = part.get("cost_centre")
    status = get_cost_centre_status(cost_centre)
    
    print(f"{name} - {cost_centre} - {status}")
    
    if status == "Valid":
        valid_count += 1
    elif status == "Invalid":
        invalid_count += 1
    else:
        missing_count += 1

print(f"Valid count: {valid_count}")
print(f"Invalid count: {invalid_count}")
print(f"Missing count: {missing_count}")