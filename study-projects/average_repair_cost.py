jobs = [
    {"job": "J301", "machine": "Emba 2", "cost": 120, "completed": True},
    {"job": "J302", "machine": "Bobst 1", "cost": 90, "completed": True},
    {"job": "J303", "machine": "Emba 2", "cost": 280, "completed": True},
    {"job": "J304", "machine": "Bobst 1", "cost": 210, "completed": True},
    {"job": "J305", "machine": "Emba 2", "cost": 800, "completed": False},
    {"job": "J306", "machine": "Mosca 5", "cost": 75, "completed": True}
]

def build_repair_summary(jobs):
    summary = {}
    for job in jobs:
        machine = job["machine"]
        cost = job["cost"]
        completed = job["completed"]
        if completed:
            if machine not in summary:
                summary[machine] = {"count": 0, "total_cost": 0}
            summary[machine]["count"] += 1
            summary[machine]["total_cost"] += cost
    return summary


repair_summary = build_repair_summary(jobs)


if repair_summary:
    for machine in repair_summary:
        details = repair_summary[machine]
        average = details["total_cost"] / details["count"]
        print(f"{machine} - Count: {details["count"]} - Total cost: £{details["total_cost"]:.2f} - Average: £{average:.2f}")
else:
    print("No jobs")