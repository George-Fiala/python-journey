jobs = [
    {"job": "J701", "machine": "Emba 2", "waiting_days": 5, "cost": 180, "completed": False},
    {"job": "J702", "machine": "Bobst 1", "waiting_days": 8, "cost": 300, "completed": False},
    {"job": "J703", "machine": "Emba 2", "waiting_days": 9, "cost": 250, "completed": False},
    {"job": "J704", "machine": "Emba 2", "waiting_days": 9, "cost": 420, "completed": False},
    {"job": "J705", "machine": "Bobst 1", "waiting_days": 12, "cost": 600, "completed": True},
    {"job": "J706", "machine": "Bobst 1", "waiting_days": 8, "cost": 450, "completed": False},
    {"job": "J707", "machine": "Mosca 5", "waiting_days": 4, "cost": 120, "completed": False},
]


def find_priority_repair_by_machine(jobs):
    longest_wait = {}
    for job in jobs:
        machine = job["machine"]
        waiting_days = job["waiting_days"]
        cost = job["cost"]
        completed = job["completed"]
        if not completed:
            if machine not in longest_wait:
                longest_wait[machine] = job
            elif waiting_days > longest_wait[machine]["waiting_days"]:
                longest_wait[machine] = job
            elif (waiting_days == longest_wait[machine]["waiting_days"]
                  and cost > longest_wait[machine]["cost"]):
                longest_wait[machine] = job
    return longest_wait


priority_repair = find_priority_repair_by_machine(jobs)

if priority_repair:
    for machine in priority_repair:
        details = priority_repair[machine]
        print(f"{machine}: Job: {details['job']} - Waiting days: {details['waiting_days']} - Cost: £{details['cost']:.2f}")
else:
    print("No repairs")
