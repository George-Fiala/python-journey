machines_data = [
    {"machine_id": "Press_01", "leaks": 15, "cost_per_leak_gbp": 20},
    {"machine_id": "Robot_Arm_A", "leaks": 2, "cost_per_leak_gbp": 50},
    {"machine_id": "Conveyor_Belt_Main", "leaks": 8, "cost_per_leak_gbp": 10}, 
    {"machine_id": "Press_02", "leaks": 8, "cost_per_leak_gbp": 25},
    {"machine_id": "Packaging_Unit", "leaks": 20, "cost_per_leak_gbp": 15}
]

print("--- AIR LEAK AUDIT ---")

total_cost = 0
critical_machines = []
audit_report = []

for machine in machines_data:
    each_machine_cost = machine["leaks"] * machine["cost_per_leak_gbp"]

    print(f"Repair cost per machine: £{each_machine_cost}")
    total_cost += each_machine_cost
    if machine["leaks"] > 5:
        critical_machines.append(machine["machine_id"])
    
    machine_record = {
        "id": machine["machine_id"],
        "bill": each_machine_cost
    }

    audit_report.append(machine_record)


print("-" * 30)
print(f"Total cost: {total_cost}")
print(f"CRITICAL MACHINES (Fix immediately): {critical_machines}")
print("-" * 30)
print("FINAL AUDIT REPORT (Send to Manager):")
for row in audit_report:
    print(row)
