#!/usr/bin/env python3

# Script: Environment Variables Printer and Terminal Clearer
# Title: Environment Variables Printer and Terminal Clearer
# Description: A Python script to print all environment variables and clear the terminal screen.

# Author: bryan.broussard
# Email: bryanbroussard@gmail.com
# Date: 2025-12-19

# This script prints all environment variables in the current OS environment.
# It also includes a function to clear the terminal screen based on the OS.
# Supported platforms are Linux, macOS (darwin), and Windows.
# If the platform is not supported, it prints an error message and exits.

# Example output:
# {'PATH': 'C:\\Windows\\System32;C:\\Windows;C:\\Program Files\\Python39;...', ...}
# Each environment variable and its value printed on a new line.


# License: MIT
# Version: 1.0
# Tags: environment, variables, os, platform, clear, terminal
# Language: Python 3.x
# Dependencies: None (uses standard library)
# Tested on: Windows 10, Windows 11, Linux, macOS
# Note: Modify the script as needed for additional functionality.

# Import necessary modules
import os
import sys

# Determine the clear command based on the operating system
if sys.platform in ('linux', 'darwin'):
    CLEAR = 'clear'
elif sys.platform == 'win32':
    CLEAR = 'cls'
else:
    print('Platfrom not supported', file=sys.stderr)
    exit(1)

# Function to clear the terminal screen
def clear_term() -> None:
    os.system(CLEAR)
    
# Print all environment variables as a dictionary-like object
print(os.environ)

# Print each environment variable and its value on a new line
for key, value in os.environ.items():
    print(f"{key}: {value}")
