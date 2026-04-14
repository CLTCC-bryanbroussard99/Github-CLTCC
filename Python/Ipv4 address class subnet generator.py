#!/usr/bin/env python3
"""
IPv4 Address Class Generator — CSV output
Generates random IPv4 addresses and their class (A, B, C, D, E).
"""

import random
import argparse


SUBNET_MASKS = {
    "A": "255.0.0.0",
    "B": "255.255.0.0",
    "C": "255.255.255.0",
    "D": "N/A",
    "E": "N/A",
}

def get_ipv4_class(ip: str) -> str:
    first = int(ip.split(".")[0])
    if 1 <= first <= 127:
        return "A"
    elif 128 <= first <= 191:
        return "B"
    elif 192 <= first <= 223:
        return "C"
    elif 224 <= first <= 239:
        return "D"
    else:
        return "E"


def generate_random_ipv4() -> str:
    first = random.randint(1, 255)
    rest = [random.randint(0, 255) for _ in range(3)]
    return f"{first}.{rest[0]}.{rest[1]}.{rest[2]}"


def main():
    parser = argparse.ArgumentParser(
        description="Generate random IPv4 addresses and output as CSV."
    )
    parser.add_argument(
        "-n", "--count",
        type=int,
        default=20,
        help="Number of addresses to generate (default: 20)"
    )
    args = parser.parse_args()

    print("ip_address,class,subnet_mask")
    for _ in range(args.count):
        ip = generate_random_ipv4()
        cls = get_ipv4_class(ip)
        print(f"{ip},{cls},{SUBNET_MASKS[cls]}")


if __name__ == "__main__":
    main()