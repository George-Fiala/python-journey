
user_name = input("What's your name, warrior?: ")
print(f"Hi {user_name}, let's code!")   
while True:
    try:
        monthly_deposit = input("How much £ monthly to BTC?: ")
        total = int(monthly_deposit) * 12
        deposit_num = total
        if deposit_num > 12000:
            print("Whale level—massive BTC growth ahead!")
        elif deposit_num > 6000:
            print("Warrior steady—solid yearly plan!")
        else:
            print("Build up—every £ counts toward 5k £ goals!")
        
        break
    except:
        print("Invalid input enter number")
        continue
while True:
    try:
        salary_goal = input("3-year salary goal (£)? ")
        goal_num = int(salary_goal)
        if goal_num >=5000:
            print("Warrior level-crush it!")
        else:
            print("Aim higher add more BTC deposits!")
        break
    except:
        print("Invalid goal—enter number!")
        continue

print(f"With {deposit_num / 12} monthly, hit {goal_num} £ faster!")