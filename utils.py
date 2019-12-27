# Requires xlrd and openpyxl
import pandas as pd
import os
import sys
from collections import Counter

NAME = 'Name'
COLS_TO_KEEP = ['phone', 'email']   # Cols to keep (not including Name col)
RECORD_FILE = '.data.csv'           # File to collect/store data

def display_help():
    msg = """
HELP PAGE:
help: Display help page (this page)
add: Add a new Excel Spreadsheet to the collection of records
add from directory: Given a directory, find all spreadsheets and add to records
output: Create an Excel spreadsheet using all records
delete: Delete event data by event name
list events: Print a list of all events in the database
reset: Delete all records
quit: quit this program"""
    # print(msg)
    return msg


def add_spreadsheet():
    read_from = input("Name of Excel Spreadsheet containing new data: ")
    return add(read_from)


def add_from_directory():
    directory = input("Name of directory holding Excel Spreadsheets: ")
    try:
        spreadsheets = [f for f in os.listdir(directory) if f[-5:] == '.xlsx']
    except:
        # print("FAIL: Bad directory")
        return "FAIL: Bad directory"

    ret = ''
    for sheet in spreadsheets:
        ret += add(directory + sheet) + "\n"
    
    return ret


def add(sheet):
    try:
        data = pd.read_excel(sheet)
    except:
        # print('FAIL: Could not read file', sheet)
        return f'FAIL: Could not read file {str(sheet)}'

    for col_name in [NAME] + COLS_TO_KEEP:
        if col_name not in data.columns:
            # print(f'FAIL: Did not find column {col_name}')
            return f'FAIL: Did not find column {col_name}'

    # Remove multiple instances of names, keeping the first instance of each name
    new_data = data.groupby(NAME, as_index = False).first()
    new_data = new_data[[NAME] + COLS_TO_KEEP]

    if RECORD_FILE not in os.listdir():
        event_name = None
        while not event_name:
            event_name = input(f"Name of event for {sheet} (MUST BE UNIQUE): ")
        new_data[event_name] = 1
        new_data.to_csv(RECORD_FILE, index=False)
    else:
        all_data = pd.read_csv(RECORD_FILE)

        event_name = None
        while not event_name or event_name in all_data.columns:
            event_name = input("Name of event (MUST BE UNIQUE): ")
        new_data[event_name] = 1

        all_data = pd.concat([all_data, new_data], sort = False)
        all_data.to_csv(RECORD_FILE, index=False)
    # print(f"SUCCESS: Added {sheet} to records as '{event_name}'")
    return f"SUCCESS: Added {sheet} to records as '{event_name}'"


def create_excel():
    if RECORD_FILE not in os.listdir():
        # print("FAIL: No data exists (hidden file '.data.csv' not found)")
        return "FAIL: No data exists (hidden file '.data.csv' not found)"

    all_data = pd.read_csv(RECORD_FILE)
    final = all_data.groupby(NAME, as_index = False).first()

    name_count = Counter(all_data[NAME])
    final["Counts"] = final[NAME].apply(lambda x: name_count[x])

    output_file = input("Name of output file: ")
    if not output_file:
        # print("FAIL: No output file given!")
        return "FAIL: No output file given!"
    if output_file[-5:] != ".xlsx":
        output_file += ".xlsx"
    final.to_excel(output_file)
    # print(f"SUCCESS: Created file {output_file}")
    return f"SUCCESS: Created file {output_file}"


def delete_event():
    all_data = pd.read_csv(RECORD_FILE)

    event_name = None
    while not event_name or event_name not in all_data.columns:
        event_name = input( "Name of event to delete (enter 'list events' to view all events, " +
                            "'cancel' to return to main menu): ")
        if event_name == 'list events':
            list_events()
        elif event_name == 'cancel':
            return

    all_data = all_data[all_data[event_name].isna()]
    all_data = all_data.drop(columns=[event_name])
    all_data.to_csv(RECORD_FILE, index=False)

    # print(f"SUCCESS: Deleted event '{event_name}'")
    return f"SUCCESS: Deleted event '{event_name}'"


def list_events():
    if RECORD_FILE not in os.listdir():
        # print(f"FAIL: No events to list because {RECORD_FILE} not found!")
        return f"FAIL: No events to list because {RECORD_FILE} not found!"

    all_data = pd.read_csv(RECORD_FILE)
    events_list = f"\n".join([str(col) for col in all_data.columns if col not in [NAME] + COLS_TO_KEEP])
    # print("\nExisting Events:\n" + events_list + "\n")
    return "Existing Events:\n" + events_list


def reset_records():
    if input("CONFIRM (enter 'yes'): ") == 'yes':
        return reset()


def reset():
    if RECORD_FILE in os.listdir():
        os.remove(RECORD_FILE)
        return "SUCCESS: Performed reset"
    else:
        return "FAIL: No reset necessary"


def quit_program():
    sys.exit("Quitting")