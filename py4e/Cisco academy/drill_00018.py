tahak = {}

def fib_chytre(n):
    if n in tahak:
        return tahak[n]

    if n == 0:
        return 0

    elif n == 1:
        return 1

    else:
        vysledek = fib_chytre(n - 1) + fib_chytre(n - 2)
        tahak[n] = vysledek
        return vysledek


print(fib_chytre(900))
