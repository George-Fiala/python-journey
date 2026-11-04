breakdowns = [
    {"id": "B101", "machine": "Emba 2", "downtime": 45, "planned": False},
    {"id": "B102", "machine": "Bobst 1", "downtime": 20, "planned": False},
    {"id": "B103", "machine": "Emba 2", "downtime": 30, "planned": True},
    {"id": "B104", "machine": "Emba 2", "downtime": 75, "planned": False},
    {"id": "B105", "machine": "Bobst 1", "downtime": 40, "planned": False},
    {"id": "B106", "machine": "Gopfert 1", "downtime": 90, "planned": True},
    {"id": "B107", "machine": "Mosca 5", "downtime": 15, "planned": False}
]



def build_breakdown_summary(breakdowns):
    summary = {}
    for breakdown in breakdowns:
        machine = breakdown["machine"]
        downtime = breakdown["downtime"]
        planned = breakdown["planned"]
        if not planned:
            if machine not in summary:
                summary[machine] = {"count": 0, "total_downtime": 0}
            summary[machine]["count"] += 1
            summary[machine]["total_downtime"] += downtime
    return summary


breakdown_summary = build_breakdown_summary(breakdowns)

if breakdown_summary:
    for machine in breakdown_summary:
        details = breakdown_summary[machine]
        print(f"{machine} - count: {details["count"]} - total downtime: {details["total_downtime"]}")

else:
    print("No downtime")