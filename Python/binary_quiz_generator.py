import html

filename = "binary_quiz.txt"
start = 0
end = 255

def binary_quiz_generator(start, end):
        
    with open(filename, "w", encoding="utf-8") as file:

        file.write("Quiz title: Subnetting\nQuiz description: Learn.\n\n")  # add spacing between blocks

        for decimal in range(start, end + 1):

            # Format as 8-bit padded binary
            binary_val = format(decimal, '08b')
            
            # Formatted HTML Block
            raw_html = f"""{decimal}. <div style="width: 99%; margin: 0 auto; font-family: 'Segoe UI', Arial, sans-serif; background-color: #ffffff; color: #2d3436; line-height: 1.6; padding: 20px; border: 1px solid #dfe6e9; border-radius: 12px;"><header style="background-color: #d63031; color: #ffffff; padding: 20px; border-radius: 8px; margin-bottom: 25px;"><h2 style="margin: 0; font-size: 1.5rem;">Knowledge Check: Binary Conversion</h2><p style="margin: 5px 0 0 0;">Skill Assessment</p></header><section style="background-color: #f8f9fa; border-left: 6px solid #d63031; padding: 25px; border-radius: 4px; margin-bottom: 20px;"><p style="font-size: 1.2rem; margin-bottom: 20px;">Using the positional addition method for 8-bit octets, convert the following <strong>Binary</strong> value into its <strong>Decimal</strong> equivalent:</p><div style="display: inline-block; background-color: #2d3436; color: #00ff00; font-family: 'Courier New', Courier, monospace; font-size: 2rem; padding: 15px 30px; border-radius: 6px; margin-bottom: 20px;">{binary_val}</div><br /><strong style="background-color: #ffffff; font-size: 1rem;"</strong></section><footer style="margin-top: 30px; padding: 15px; background-color: #e1f5fe; border-radius: 6px; border: 1px solid #b3e5fc;"><h3 style="margin: 0 0 10px 0; color: #01579b;">Instructor's Guidance:</h3><p style="margin: 0; font-size: 0.95rem;">To solve this, align the binary string to your 8-bit grid (128, 64, 32, 16, 8, 4, 2, 1) and sum the values where a '1' is present.</p></footer></div>\n* {decimal}"""

            file.write(raw_html + "\n\n")  # add spacing between blocks

    print(f"Success: Quiz blocks for {start}-{end} saved to {filename}")

if __name__ == "__main__":
    binary_quiz_generator(start, end)