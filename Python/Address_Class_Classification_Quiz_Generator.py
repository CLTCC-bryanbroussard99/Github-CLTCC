#!/usr/bin/env python3
"""
IPv4 Address Class Generator
Generates random IPv4 addresses and identifies their class (A, B, C, D, E).
"""

import random
import ipaddress
import argparse


def get_ipv4_class(ip: str) -> tuple[str, str]:
    """
    Determine the class of an IPv4 address based on the first octet.

    Returns:
        (class_letter, description)
    """
    first_octet = int(ip.split(".")[0])

    if 1 <= first_octet <= 127:
        return "A", "Large networks"
    elif 128 <= first_octet <= 191:
        return "B", "Medium networks"
    elif 192 <= first_octet <= 223:
        return "C", "Small networks"
    elif 224 <= first_octet <= 239:
        return "D", "Multicast"
    elif 240 <= first_octet <= 255:
        return "E", "Experimental / Reserved"
    else:
        return "?", "Unknown"


def is_private(ip: str) -> bool:
    """Check if the IP address is in a private range."""
    return ipaddress.ip_address(ip).is_private


def get_default_subnet(ip_class: str) -> str:
    """Return the default subnet mask for a given IP class."""
    masks = {
        "A":  "255.0.0.0     (/8)",
        "A*": "255.0.0.0     (/8)",
        "B":  "255.255.0.0   (/16)",
        "C":  "255.255.255.0 (/24)",
        "D":  "N/A (Multicast)",
        "E":  "N/A (Experimental)",
        "?":  "N/A",
    }
    return masks.get(ip_class, "N/A")


def generate_random_ipv4() -> str:
    """Generate a random valid IPv4 address (excluding 0.x.x.x)."""
    first = random.randint(1, 255)
    rest = [random.randint(0, 255) for _ in range(3)]
    return f"{first}.{rest[0]}.{rest[1]}.{rest[2]}"


def generate_addresses(count: int, show_private: bool = True) -> list[dict]:
    """Generate a list of IPv4 address info dictionaries."""
    results = []
    for _ in range(count):
        ip = generate_random_ipv4()
        ip_class, description = get_ipv4_class(ip)
        private = is_private(ip)
        subnet = get_default_subnet(ip_class)
        results.append({
            "ip":          ip,
            "class":       ip_class,
            "description": description,
            "private":     private,
            "subnet":      subnet,
        })
    return results


def print_table(entries: list[dict]) -> None:
    """Print results in a formatted table."""
    header = f"{'IP Address':<18} {'Class':<6} {'Private':<9} {'Default Subnet':<26} {'Description'}"
    divider = "-" * 85
    print(divider)
    print(header)
    print(divider)
    for e in entries:
        private_flag = "Yes" if e["private"] else "No"
        print(
            f"{e['ip']:<18} {e['class']:<6}  "
        )
    print(divider)


def print_class_summary(entries: list[dict]) -> None:
    """Print a summary count by class."""
    from collections import Counter
    counts = Counter(e["class"] for e in entries)
    print("\nClass Distribution:")
    for cls in sorted(counts):
        print(f"  Class {cls}: {counts[cls]}")
    print()


def main():
    parser = argparse.ArgumentParser(
        description="Generate random IPv4 addresses and display their class info."
    )
    parser.add_argument(
        "-n", "--count",
        type=int,
        default=255,
        help="Number of IP addresses to generate (default: 20)"
    )
    parser.add_argument(
        "--summary",
        action="store_true",
        help="Show class distribution summary"
    )
    args = parser.parse_args()

    print(f"\nGenerating {args.count} random IPv4 addresses...\n")
    entries = generate_addresses(args.count)
    print_table(entries)

    if args.summary:
        print_class_summary(entries)

    print("Class Reference:")
    print("  A  (1–126)   — Large networks,  default mask /8")
    print("  A* (127)     — Loopback reserved")
    print("  B  (128–191) — Medium networks, default mask /16")
    print("  C  (192–223) — Small networks,  default mask /24")
    print("  D  (224–239) — Multicast")
    print("  E  (240–255) — Experimental / Reserved\n")


if __name__ == "__main__":
    main()