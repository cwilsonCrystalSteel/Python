# -*- coding: utf-8 -*-
"""
Created on Tue Mar 11 17:59:07 2025

@author: Netadmin
"""
import os
from pathlib import Path

# Define the new config file path
config_file_path = r"set CONFIG_FILE=C:\Users\Netadmin\Documents\GitHub\Python\batFileLocations.txt"

# Get the current working directory
cwd = Path(os.getcwd())

# Find all .bat files in the directory and subdirectories
bat_files = list(cwd.rglob("*.bat"))

# Loop through each .bat file and update the line
for bat_file in bat_files:
    try:
        with open(bat_file, "r", encoding="utf-8") as file:
            lines = file.readlines()

        # Update the specific line if found
        updated_lines = [
            config_file_path + "\n" if line.startswith("set CONFIG_FILE=") else line
            for line in lines
        ]

        # Write the updated content back to the file
        with open(bat_file, "w", encoding="utf-8") as file:
            file.writelines(updated_lines)

        print(f"Updated: {bat_file}")

    except Exception as e:
        print(f"Error processing {bat_file}: {e}")
