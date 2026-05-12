#!/usr/bin/env python3
"""
IPv4 Subnet Generator — CSV output
Generates IPv4 addresses and detailed subnetting parameters.
"""
import random
import argparse

DEFAULT_MASKS = {
    "A": "255.0.0.0",
    "B": "255.255.0.0",
    "C": "255.255.255.0",
}

def get_ipv4_class(ip: str) -> str:
    first = int(ip.split(".")[0])
    if 1 <= first <= 127: return "A"
    if 128 <= first <= 191: return "B"
    if 192 <= first <= 223: return "C"
    if 224 <= first <= 239: return "D"
    return "E"

def cidr_to_netmask(bits: int) -> str:
    mask = (0xffffffff << (32 - bits)) & 0xffffffff
    return f"{(mask >> 24) & 0xff}.{(mask >> 16) & 0xff}.{(mask >> 8) & 0xff}.{mask & 0xff}"

def generate_random_ipv4() -> str:
    # Avoiding 127.x.x.x (loopback) for cleaner subnetting examples
    first = random.choice([r for r in range(1, 224) if r != 127])
    return f"{first}.{random.randint(0,255)}.{random.randint(0,255)}.{random.randint(0,255)}"

def main():
    parser = argparse.ArgumentParser(description="Generate IPv4 subnetting data as CSV.")
    parser.add_argument("-n", "--count", type=int, default=20, help="Number of addresses")
    args = parser.parse_args()

    # Header as requested
    print("Address,Class,Default Subnet Mask,Custom Subnet Mask,Total Subnets,Total Host Addresses,Usable Addresses,Bits Borrowed")

    for _ in range(args.count):
        ip = generate_random_ipv4()
        cls = get_ipv4_class(ip)
        
        if cls in DEFAULT_MASKS:
            default_bits = {"A": 8, "B": 16, "C": 24}[cls]
            # Randomly borrow between 1 and (max bits - 2) to keep usable hosts
            bits_borrowed = random.randint(1, 30 - default_bits)
            custom_bits = default_bits + bits_borrowed
            
            custom_mask = cidr_to_netmask(custom_bits)
            total_subnets = 2**bits_borrowed
            total_hosts = 2**(32 - custom_bits)
            usable_hosts = max(0, total_hosts - 2)
            
            print(f"{ip},{cls},{DEFAULT_MASKS[cls]},{custom_mask},{total_subnets},{total_hosts},{usable_hosts},{bits_borrowed}")
        else:
            # Classes D/E do not follow standard subnetting
            print(f"{ip},{cls},N/A,N/A,N/A,N/A,N/A,N/A")

if __name__ == "__main__":
    main()