machines = [
    {"name": "Emba 1", "downtime": 25, "planned": False},
    {"name": "Gopfert 1", "downtime": 90, "planned": False},
    {"name": "Bobst 2", "downtime": 45, "planned": True},
    {"name": "Asahi 1", "downtime": 10, "planned": False},
]


def get_downtime_status(downtime, planned):
    if planned:
        return "Planned"
    elif downtime >= 60:
        return "Critical"
    elif downtime >= 20:
        return "Warning"
    else:
        return "Normal"


total_downtime = 0
critical_machines_count = 0
unplanned_downtime_total = 0

for machine in machines:
    name = machine["name"]
    downtime = machine["downtime"]
    planned = machine["planned"]
    downtime_status = get_downtime_status(downtime, planned)
    total_downtime += downtime
    if downtime_status == "Critical":
        critical_machines_count += 1

    if not planned:
        unplanned_downtime_total += downtime

    print(f"{name} - Downtime: {downtime} - Status: {downtime_status}")

    

print(f"Total downtime: {total_downtime}")
print(f"Total unplanned downtime: {unplanned_downtime_total}")
print(f"Critical machines count: {critical_machines_count}")



    