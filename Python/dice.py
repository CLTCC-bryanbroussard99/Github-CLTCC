import random

# Available dice
dice_sides = [100, 20, 12, 6, 4]

def roll_die(sides, times=1):
    """Roll a die 'times' times and return a list of results."""
    return [random.randint(1, sides) for _ in range(times)]

def main():
    print("Welcome to the Dice Roller!")
    print(f"Available dice: {dice_sides}")
    
    while True:
        choice = input("\nEnter which die to roll (e.g., 20) or 'q' to quit: ").strip().lower()
        if choice == 'q':
            print("Goodbye!")
            break
        
        try:
            sides = int(choice)
            if sides not in dice_sides:
                print(f"Invalid choice. Please choose from {dice_sides}.")
                continue
            
            times = input(f"How many {sides}-sided dice do you want to roll? ").strip()
            times = int(times)
            if times <= 0:
                print("Number of dice must be at least 1.")
                continue
            
            results = roll_die(sides, times)
            print(f"Rolling {times} {sides}-sided dice: {results}")
            print('-----------------------------')
            print(f"Total: {sum(results)}\nMin: {min(results)}\nMax: {max(results)}\nAverage: {sum(results)/times:.2f}")
            print('-----------------------------')
        
        except ValueError:
            print("Please enter valid numbers.")
        
if __name__ == "__main__":
    main()