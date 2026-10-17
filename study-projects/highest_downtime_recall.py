machines = [
    {"machine": "Emba 1", "downtime": 35, "planned": False, "repair_cost": 300},
    {"machine": "Bobst 2", "downtime": 80, "planned": True, "repair_cost": 7000},
    {"machine": "Gopfert 1", "downtime": 65, "planned": False, "repair_cost": 1500},
    {"machine": "Emba 4", "downtime": 65,  "planned": False, "repair_cost": 700}
]


def find_highest_downtime(machines):
    highest_downtime = None
    for machine in machines:
        downtime = machine["downtime"]
        planned = machine["planned"]
        repair_cost = machine["repair_cost"]
        if not planned:
            if highest_downtime is None or highest_downtime["downtime"] < downtime:  
                highest_downtime = machine

            elif highest_downtime["downtime"] == downtime  and highest_downtime["repair_cost"] < repair_cost:
                highest_downtime = machine
        
            
                
           
                    
            
    return highest_downtime



max_downtime = find_highest_downtime(machines)
if max_downtime is not None:
    machine = max_downtime["machine"]
    downtime = max_downtime["downtime"]
    repair_cost = max_downtime["repair_cost"]
    print(f"{machine} - {downtime}min - £{repair_cost}")
else:
    print("No unplanned downtime")