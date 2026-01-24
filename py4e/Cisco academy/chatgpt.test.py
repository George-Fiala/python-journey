efficiencies = [0.82, 0.76, 0.91, 0.0, 0.67, 0.88, 0.93]

total_efficiency = 0
best_efficiency = None
worst_efficiency = None
average_all = 0
average_running = 0
count_zero = 0
count_above_target = 0
best_index = 0
running = 0
running_count = 0
index = 0

for e in efficiencies:
    total_efficiency += e
    if best_efficiency is None or e > best_efficiency:
        best_efficiency = e
        best_index = index
    if worst_efficiency is None or e < worst_efficiency:
        worst_efficiency = e
    index += 1   
    
    
    
    if e > 0:
        running += e
        running_count += 1
        average_running =running / running_count
    elif e < 0:
        average_running == None
    else:
        average_running == None
        
    if e == 0:
        count_zero += 1

    if e >= 0.85:
        count_above_target += 1

average_all = total_efficiency / len(efficiencies)

print("Total efficiency:", total_efficiency)
print("Average:", average_all)
print("Average running:", average_running)
print("Best efficiency:", best_efficiency)
print("Worst efficiency:", worst_efficiency)
print("Zero days:", count_zero)
print("Shifts above targert:", count_above_target)
print("Total shifts:", len(efficiencies))
print("Bext index:", best_index)
