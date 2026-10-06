machines = [
    {"machine_id": "MO1001"},
    {"machine_id": "MO1002"},
    {"machine_id": "MO1003"},
    {"machine_id": "MO1002"},
    {"machine_id": "MO1004"},
    {"machine_id": "MO1001"},
    {"machine_id": "MO1002"}
]


seen = []
duplicates = []


for machine in machines:
    machine_id = machine.get("machine_id")
    if machine_id in seen:
        if machine_id not in duplicates:
            duplicates.append(machine_id)
    else:
        seen.append(machine_id)


print(f"Seen: {seen}")
print(f"Duplicates: {duplicates}")