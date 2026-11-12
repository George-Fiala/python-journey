jobs = [
    {"job": "J601", "machine": "Emba 2", "cost": 180, "completed": False},
    {"job": "J602", "machine": "Emba 2", "cost": 440, "completed": False},
    {"job": "J603", "machine": "Bobst 1", "cost": 300, "completed": False},
    {"job": "J604", "machine": "Bobst 1", "cost": 500, "completed": True},
    {"job": "J605", "machine": "Emba 2", "cost": 220, "completed": False},
    {"job": "J606", "machine": "Mosca 5", "cost": 150, "completed": False},
]


def find_highest_repair_by_machine(jobs):
    most_expensive = {}
    for job in jobs:
        machine = job["machine"]
        cost = job["cost"]
        completed = job["completed"]
        if not completed:
            if machine not in most_expensive:
                most_expensive[machine] = job
            elif cost > most_expensive[machine]["cost"]:
                most_expensive[machine] = job
    return most_expensive


repair_by_machine = find_highest_repair_by_machine(jobs)

if repair_by_machine:
    for machine in repair_by_machine:
        details = repair_by_machine[machine]
        print(f"{machine} - {details['job']} - cost: £{details['cost']:.2f}")
else:
    print("No repairs")