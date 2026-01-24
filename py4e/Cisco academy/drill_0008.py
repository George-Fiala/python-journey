hour = int(input("Starting time (hours): "))
mins = int(input("Starting time (minutes): "))
dura = int(input("Event duration (minutes): "))
total_min = mins + dura# find a total of all minutes
final_hour = (total_min // 60) + hour # find a number of hours hidden in minutes and update the hour
min = total_min % 60 # correct minutes to fall in the (0..59) range
h = ((total_min // 60) + hour) % 24 # correct hours to fall in the (0..23) range
m = total_min - 60

print(h, ":", min, sep=' ')




