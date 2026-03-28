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


def plot_prime_count(n, primes, output_path="primes_chart.png"):
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt

    prime_set = set(primes)
    x = list(range(n))
    y = []
    count = 0
    for i in x:
        if i in prime_set:
            count += 1
        y.append(count)

    plt.figure(figsize=(10, 5))
    plt.plot(x, y, color="steelblue", linewidth=1.5)
    plt.xlabel("Integer")
    plt.ylabel("Number of Primes")
    plt.title(f"Prime Counting Function π(x) for x < {n}")
    plt.grid(True, alpha=0.3)
    plt.tight_layout()

    plt.savefig(output_path, dpi=150)
    plt.close()
    print(f"Chart saved to: {output_path}")


def main():
    import sys
    if len(sys.argv) > 1:
        n = int(sys.argv[1])
    else:
        n = int(input("Enter a number: "))
    primes = get_primes(n)
    if primes:
        print(f"Primes less than {n}:")
        print(primes)
    else:
        print(f"No primes less than {n}.")
    plot_prime_count(n, primes)


if __name__ == "__main__":
    main()
