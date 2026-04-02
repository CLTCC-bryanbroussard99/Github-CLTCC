import csv

def process_ipv4_data(input_file, output_file):
    """
    Reads network data from CSV, prints to console, and saves to a text file.
    """
    print(f"--- Processing {input_file} ---\n")
    
    try:
        with open(input_file, mode='r', encoding='utf-8') as csv_file, \
             open(output_file, mode='w', encoding='utf-8') as txt_file:
            
            reader = csv.DictReader(csv_file)
            
            # Write a title to the quiz file
            txt_file.write("Quiz title: IP Address Decomposition Lab\n\n")
            
            for i, row in enumerate(reader, 1):
                ip = row.get('ipv4 address', 'N/A')
                net = row.get('network portion', 'N/A')
                host = row.get('host portion', 'N/A')
                ip_class = row.get('class', 'N/A')
                
                output_line = f"""{i}. <div style="width: 99%; margin: 20px auto; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; color: #333;"><header style="margin-bottom: 25px; border-left: 5px solid #2c3e50; padding-left: 15px;"><h2 style="margin: 0; color: #2c3e50; font-size: 1.5rem;">Lab Exercise: IPv4 Classful Analysis</h2><p style="margin: 5px 0 0; color: #666;">Analyze the following target address and identify its structural components.</p></header><div style="background: #ffffff; border: 1px solid #ddd; border-radius: 8px; overflow: hidden;" role="region" aria-labelledby="target-address-heading"><div style="background: #f8f9fa; padding: 20px; border-bottom: 1px solid #eee; text-align: center;"><span id="target-address-heading" style="display: block; font-size: 0.9rem; color: #666; margin-bottom: 8px;">Target IPv4 Address</span><div style="font-family: 'Courier New', Courier, monospace; font-size: 2.5rem; color: #d35400;">{ip}</div></div><div style="padding: 25px;"><div style="margin-bottom: 20px;">1. Address Class<div style="font-size: 0.85rem; color: #555; margin-bottom: 8px;">Identify if this is Class A, B, or C based on the first octet.</div>Class (A, B, C , D, E): \n* [{ip_class}]</div><div style="margin-bottom: 20px;">2. Network Portion (Prefix)<div style="font-size: 0.85rem; color: #555; margin-bottom: 8px;">Enter the octets representing the Network ID (e.g., X.X.X).&nbsp;</div><div style="font-size: 0.85rem; color: #555; margin-bottom: 8px;">Network: \n* [{net}]</div></div><div style="margin-bottom: 25px;">3. Host Portion<div style="font-size: 0.85rem; color: #555; margin-bottom: 8px;">Enter the remaining octets representing the specific host interface.</div><div style="font-size: 0.85rem; color: #555; margin-bottom: 8px;">Host: \n* [{host}]</div></div></div></div></div>"""

                # Output to text file
                txt_file.write(output_line + "\n")
                
        print(f"\nSuccess: Results saved to {output_file}")

    except FileNotFoundError:
        print(f"Error: The file '{input_file}' was not found.")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

# Configuration
if __name__ == "__main__":
    FILE_IN = "randomized_ipv4_lab.csv"
    FILE_OUT = "NetHostClass_quiz.txt"
    
    process_ipv4_data(FILE_IN, FILE_OUT)