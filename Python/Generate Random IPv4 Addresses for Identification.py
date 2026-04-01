import csv
import random
import ipaddress

def generate_random_classful_ips(count_per_class=75):
    results = []
    
    # Define the ranges for Class A, B, and C
    # We avoid 127.x.x.x (Loopback) and 0.x.x.x
    ranges = [
        ('A', 1, 126),
        ('B', 128, 191),
        ('C', 192, 223)
    ]

    for cls_name, start, end in ranges:
        found = 0
        while found < count_per_class:
            # Generate a random first octet in the specific class range
            first = random.randint(start, end)
            rest = ".".join(str(random.randint(0, 255)) for _ in range(3))
            ip_str = f"{first}.{rest}"
            
            octets = ip_str.split('.')
            
            if cls_name == 'A':
                net, host = octets[0], ".".join(octets[1:])
            elif cls_name == 'B':
                net, host = ".".join(octets[:2]), ".".join(octets[2:])
            else: # Class C
                net, host = ".".join(octets[:3]), octets[3]
                
            results.append([ip_str, net, host, f"{cls_name}"])
            found += 1
            
    # Shuffle the entire list so classes aren't grouped together
    random.shuffle(results)
    return results

# Execution and CSV Generation
data = generate_random_classful_ips(count_per_class=75) # Totals 225 addresses

with open('randomized_ipv4_lab.csv', 'w', newline='') as f:
    writer = csv.writer(f)
    writer.writerow(['ipv4 address', 'network portion', 'host portion', 'class'])
    writer.writerows(data)

print(f"Successfully generated {len(data)} randomized addresses.")