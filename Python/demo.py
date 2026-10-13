def odd_numbers(maximum):
    return_string = ""
    for num in range(1, maximum + 1, 2):
        return_string += str(num) + " "
    return return_string.strip()

print(odd_numbers(6))
print(odd_numbers(10))
print(odd_numbers(1))
print(odd_numbers(3))
print(odd_numbers(0))