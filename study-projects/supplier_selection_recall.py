quotes = [
    {"supplier": "NORD", "price": 450, "delivery": 5, "approved": True},
    {"supplier": "ERIKS", "price": 390, "delivery": 8, "approved": True},
    {"supplier": "Rubix", "price": 420, "delivery": 3, "approved": False},
    {"supplier": "BearingCo", "price": 450, "delivery": 2, "approved": True},
]


def find_best_supplier(quotes):
    cheapest_quote = None
    for quote in quotes:
        approved = quote["approved"]
        price = quote["price"]
        delivery = quote["delivery"]    
        if approved is True:
            if cheapest_quote is None or cheapest_quote["price"] > price:
                cheapest_quote = quote
            elif cheapest_quote["price"] == price and cheapest_quote["delivery"] > delivery:
                cheapest_quote = quote
    return cheapest_quote



winner = find_best_supplier(quotes)


if winner is not None:
    supplier = winner["supplier"]
    price = winner["price"]
    delivery = winner["delivery"]
    print(f"{supplier} - £{price:.2f} - {delivery}")

else:
    print("No supplier")


