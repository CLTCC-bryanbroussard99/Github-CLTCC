# Importing necessary libraries for file handling and CSV operations
# Importing the html module to ensure that any special characters in the quiz content are properly escaped when writing to the file.
import html
# The csv module is imported to facilitate reading from the CSV file containing the randomized IPv4 addresses, although in this specific implementation, we are generating quiz content based on a range of decimal values rather than directly from the CSV file.
import csv

# The filename variable is defined to specify the name of the output file where the generated quiz content will be saved. In this case, it is set to "NetHostClass_quiz.txt".
filename = "NetHostClass_quiz.txt"
# The fileimported variable is defined to specify the name of the CSV file that contains the randomized IPv4 addresses. In this case, it is set to "randomized_ipv4_lab.csv".
file_path = "randomized_ipv4_lab.csv"

def read_csv_standard(file_path):
    print(f"--- Reading {file_path} using csv module ---")
    try:
        with open(file_path, mode='r', encoding='utf-8') as file:
            # Using DictReader allows you to access columns by name
            reader = csv.DictReader(file)
            for row in reader:
                # Example: Accessing columns from your previous IP lab file
                ip = row.get('ipv4 address')
                net = row.get('network portion')
                host = row.get('host portion')
                classs = row.get('class')
                print(f"IP: {ip} | Network: {net} | Host: {host} | Class: {classs}")
    except FileNotFoundError:
        print("Error: File not found.")

# Usage
read_csv_standard('randomized_ipv4_lab.csv')