machines = [
    {"machine_id": "M001"},
    {"machine_id": "M002"},
    {"machine_id": None},
    {"machine_id": "M003"},
    {"machine_id": "M002"},
    {},    
    {"machine_id": "M001"},
    {"machine_id": None},
    
]


seen = []
duplicates = []
missing_count = 0


for machine in machines:
    machine_id = machine.get("machine_id")
    if machine_id is not None:
        if machine_id in seen:
            if machine_id not in duplicates:
                duplicates.append(machine_id)
        else:
            seen.append(machine_id)
    else:
        missing_count += 1

print(f"Seen: {seen}")
print(f"Duplicates: {duplicates}")
print(f"Count of missing: {missing_count}")