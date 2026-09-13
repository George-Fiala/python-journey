part = {"price": 420,"critical": True}

def check_part(price, critical):
    if price >= 500 and critical: 
        return "Priority"
    return "Normal"


price = part["price"]
critical = part["critical"]
status = check_part(price, critical)
print(status)