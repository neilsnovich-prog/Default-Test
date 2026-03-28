def get_primes(n):
    if n < 2:
        return []
    sieve = [True] * n
    sieve[0] = sieve[1] = False
    for i in range(2, int(n ** 0.5) + 1):
        if sieve[i]:
            for j in range(i * i, n, i):
                sieve[j] = False
    return [i for i, is_prime in enumerate(sieve) if is_prime]


def main():
    n = int(input("Enter a number: "))
    primes = get_primes(n)
    if primes:
        print(f"Primes less than {n}:")
        print(primes)
    else:
        print(f"No primes less than {n}.")


if __name__ == "__main__":
    main()
