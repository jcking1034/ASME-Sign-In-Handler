#!/usr/bin/env python3
# Requires xlrd and openpyxl
import pandas as pd
import os
from collections import Counter

NAME = 'Name'
RECORD_FILE = '.data.csv'

def display_help():
    msg = """
HELP PAGE:
help: Display help page (this page)
add: Add a new Excel Spreadsheet to the collection of records
add from directory: Given a directory, find all spreadsheets and add to records
output: Create an Excel spreadsheet using all records
reset: Delete all records
quit: quit this program"""
    print(msg)


def add_spreadsheet():
    read_from = input("Name of Excel Spreadsheet containing new data: ")
    add(read_from)


def add_from_directory():
    directory = input("Name of directory holding Excel Spreadsheets: ")
    try:
        spreadsheets = [f for f in os.listdir(directory) if f[-5:] == '.xlsx']
    except:
        print("FAIL: Bad directory")
        return

    for sheet in spreadsheets:
        add(directory + sheet)


def add(sheet):
    try:
        data = pd.read_excel(sheet)
    except:
        print('FAIL: Could not read file', sheet)
        return

    # Remove multiple instances of names, keeping the first instance of each name
    new_data = data.groupby(NAME, as_index = False).first()

    if RECORD_FILE not in os.listdir():
        new_data.to_csv(RECORD_FILE, index=False)
    else:
        all_data = pd.read_csv(RECORD_FILE)
        all_data = pd.concat([all_data, new_data], sort = False)
        all_data.to_csv(RECORD_FILE, index=False)
    print(f"SUCCESS: Added {sheet} to records")


def create_excel():
    if RECORD_FILE not in os.listdir():
        print("FAIL: No data exists (hidden file '.data.csv' not found)")
        return

    all_data = pd.read_csv(RECORD_FILE)
    final = all_data.groupby(NAME, as_index = False).first()

    name_count = Counter(all_data[NAME])
    final["Counts"] = final[NAME].apply(lambda x: name_count[x])

    output_file = input("Name of output file: ")
    if not output_file:
        print("FAIL: No output file given!")
        return
    if output_file[-5:] != ".xlsx":
        output_file += ".xlsx"
    final.to_excel(output_file)
    print(f"SUCCESS: Created file {output_file}")


def reset_records():
    if input("CONFIRM (enter 'yes'): ") == 'yes':
        if RECORD_FILE in os.listdir():
            os.remove(RECORD_FILE)
            print("SUCCESS: Performed reset")
        else:
            print("FAIL: No reset necessary")


def quit_program():
    exit()


command_list = {
    "help": display_help,
    "add": add_spreadsheet,
    "add from directory": add_from_directory,
    "output": create_excel,
    "reset": reset_records,
    "quit": quit_program
}


if __name__ == '__main__':
    while True:
        command = input("\nEnter Command ('help' for help): ")
        if command in command_list:
            print(f"\n*****\n{command}")
            command_list[command]()
        else:
            print(f"UNRECOGNIZED COMMAND: {command}")