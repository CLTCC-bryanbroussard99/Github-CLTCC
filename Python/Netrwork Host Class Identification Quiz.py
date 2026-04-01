import csv

def process_ipv4_data(input_file, output_file):
    """
    Reads network data from CSV, prints to console, and saves to a text file.
    """
    print(f"--- Processing {input_file} ---")
    
    try:
        with open(input_file, mode='r', encoding='utf-8') as csv_file, \
             open(output_file, mode='w', encoding='utf-8') as txt_file:
            
            reader = csv.DictReader(csv_file)
            
            # Write a header to the text file
            txt_file.write(f"Quiz Data Export\n{'='*50}\n")
            
            for row in reader:
                ip = row.get('ipv4 address', 'N/A')
                net = row.get('network portion', 'N/A')
                host = row.get('host portion', 'N/A')
                ip_class = row.get('class', 'N/A')
                
                output_line = f"""<div style="width: 99%; margin: 20px auto; font-family: 'Segoe UI', Arial, sans-serif; color: #1a1a1a; line-height: 1.6;"><header style="margin-bottom: 20px; border-left: 6px solid #004a99; padding-left: 15px;"><h2 style="margin: 0; color: #004a99; font-size: 1.6rem;">Cybersecurity Lab: IP Address Decomposition</h2><p style="margin: 4px 0 0; color: #444; font-weight: 500;">Objective: Determine the classful network boundaries for the target interface.</p></header><div style="background: #ffffff; border: 1px solid #d1d1d1; border-radius: 12px; box-shadow: 0 10px 20px rgba(0,0,0,0.08); overflow: hidden;" role="main" aria-label="IPv4 Assessment Card"><div style="background: #f0f4f8; padding: 30px; border-bottom: 1px solid #e1e8ed; text-align: center;"><span id="label-target-ip" style="display: block; font-size: 0.85rem; font-weight: bold; text-transform: uppercase; letter-spacing: 1.5px; color: #555; margin-bottom: 10px;">Target IPv4 Address</span><div style="font-family: 'Consolas', 'Monaco', monospace; font-size: 3rem; font-weight: 800; color: #e67e22; word-break: break-all;" aria-labelledby="label-target-ip">{ip:<15}</div></div><div style="padding: 30px;"><div style="margin-bottom: 24px;">1. Determine the Address Class<p style="margin: 0 0 12px; font-size: 0.95rem; color: #666;">Based on the first octet, which legacy class does this address belong to?{ip_class}</p>-- Choose Class </div><div style="margin-bottom: 24px;">2. Identify the Network Portion<p style="margin: 0 0 12px; font-size: 0.95rem; color: #666;">Enter the octets that constitute the Network ID (e.g., 192.168.1).</p>{net:<15}</div>argin: 0 0 12px; font-size: 0.95rem; color: #666;">Enter the octets that constitute the Network ID (e.g., 192.168.1).</p></div><div style="margin-bottom: 30px;">3. Identify the Host Portion<p style="margin: 0 0 12px; font-size: 0.95rem; color: #666;">Enter the specific octet(s) assigned to the host interface.</p></div>{host:<10}<div style="text-align: right;">Validate Analysis</div></div><div style="background: #fff9e6; border-top: 1px dashed #f39c12; padding: 20px;"><p style="margin: 0; font-size: 0.9rem; color: #7f8c8d; font-style: italic;"><strong>Instructor Reference:</strong> Class: {ip_class} | Network ID: {net} | Host ID: {host}</p></div></div></div>"""
                
                # Output to console
                print(output_line)
                
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