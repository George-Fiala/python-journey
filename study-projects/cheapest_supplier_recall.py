quotes = [
    {"supplier": "NORD", "price": 450},
    {"supplier": "ERIKS", "price": 360},
    {"supplier": "Rubix", "price": 420},
    {"supplier": "BearingCo", "price": 360}
]

def find_cheapest_supplier(quotes):
    cheapest_sup = None
    for quote in quotes:
        price = quote["price"]
        if cheapest_sup is None or cheapest_sup["price"] > price:
            cheapest_sup = quote 
    return cheapest_sup


winner = find_cheapest_supplier(quotes)


    
if winner is not None:
    supplier = winner["supplier"]
    price = winner["price"]
    print(f"{supplier} - £{price:.2f}")
else:
    print("No supplier")