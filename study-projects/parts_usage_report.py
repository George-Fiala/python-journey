parts = [
    {"part_id": "P101", "machine": "Emba 4", "cost": 160, "used": True},
    {"part_id": "P102", "machine": "Bobst 1", "cost": 300, "used": True},
    {"part_id": "P103", "machine": "Emba 4", "cost": 60, "used": False},
    {"part_id": "P104", "machine": "Emba 4", "cost": 75, "used": True},
    {"part_id": "P105", "machine": "Bobst 1", "cost": 40, "used": True},
    {"part_id": "P106", "machine": "Gopfert 1", "cost": 900, "used": False},
    {"part_id": "P107", "machine": "Mosca 5", "cost": 90, "used": True}
]


def build_used_parts_summary(parts):
    used_parts = {}
    for part in parts:
        machine = part["machine"]
        cost = part["cost"]
        used = part["used"]
        if used:
            if machine not in used_parts:
                used_parts[machine] = {"count": 0, "total_value": 0}
            used_parts[machine]["count"] += 1
            used_parts[machine]["total_value"] += cost
    return used_parts


used_parts_summary = build_used_parts_summary(parts)

if used_parts_summary:
    for machine in used_parts_summary:
        details = used_parts_summary[machine]
        print(f"{machine} - count: {details["count"]} - £{details["total_value"]:.2f}")
else:
    print("No used parts")