def is_year_leap(year):
    if year % 400 == 0:
        return True
    elif year % 100 == 0:
        return False
    elif year % 4 == 0:
        return True
    else:
        return False

def days_in_month(year, month):
    days = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    if month == 2:
        if is_year_leap(year):
            return 29
        else:
            return 28
    
    return days[month - 1]

def day_of_year(year, month, day):
    total = 0
    for m in range(1, month):
        total += days_in_month(year, m)
    return total + day
    

    
print(day_of_year(2000, 12, 31))  # 366 (přestupný rok)
print(day_of_year(1900, 3, 1))