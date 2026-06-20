orders = []

def show_menu():
    print("=== Home Study Project ===")
    print("1. Create order")
    print("2. Show orders")
    print("3. Search orders")
    print("4. Exit")

def get_choice():
    choice = input("Choose option: ")
    return choice

def create_order():
    supplier_name = input("Input supplier name: ")
    part_name = input("Input part name: ")
    quantity = input("Input quantity: ")
    cost = input("Input cost: ") 
    print("Order created for:", supplier_name,"-", part_name,"- Qty:", quantity, "-","£",cost )
    order = {
        "supplier": supplier_name,
        "part": part_name,
        "quantity": quantity,
        "cost": cost
    }
    return order
def show_orders():
    if not orders:
        print("No orders yet.")
    else:
        for order in orders:
            print("Supplier:", order["supplier"], "| Part:", order["part"], "| Qty:", order["quantity"], "| Cost: £", order["cost"])
        

def handle_choice(choice):
    if choice == "1":
        new_order = create_order()
        orders.append(new_order)     
    elif choice == "2":
        show_orders()
    elif choice == "3":
        print("Search orders selected")
    elif choice == "4":
        print("Exit selected")
    else:
        print("Invalid option")

running = True
while running:

    show_menu()
    choice = get_choice()
    print("You selected:", choice)
    handle_choice(choice)
    if choice == "4":
        running = False

 



