# Requires xlrd and openpyxl
import pandas as pd
import os
import sys
from collections import Counter


NAME = 'Name'
COLS_TO_KEEP = ['phone', 'email']   # Cols to keep (not including Name col)
RECORD_FILE = '.data.csv'           # File to collect/store data


def display_help():
    msg = \
"""\
HELP PAGE:
help: Display help page (this page)
add: Add a new Excel Spreadsheet to the collection of records
add from directory: Given a directory, find all spreadsheets and add to records
output: Create an Excel spreadsheet using all records
delete: Delete event data by event name
list events: Print a list of all events in the database
reset: Delete all records
quit: quit this program\
"""
    return msg


def add_spreadsheet():
    """
    Used by CLI ONLY
    Ask for spreadsheet and attempt to add()
    """
    read_from = input("Name of Excel Spreadsheet containing new data: ")
    if os.path.exists(read_from):
        return add(read_from)
    else:
        return "FAIL: file does not exist"


def add_from_directory():
    """
    Used by CLI ONLY
    Ask for Directory and attempt to add() all spreadsheets within it
    """
    directory = input("Name of directory holding Excel Spreadsheets: ")
    try:
        spreadsheets = [f for f in os.listdir(directory) if f[-5:] == '.xlsx']
    except:
        return "FAIL: Bad directory"

    return add_sheets(spreadsheets)


def add_sheets(sheets):
    """
    Used by GUI ONLY
    Given a list of spreadsheet names, attempt to add() each 
    """
    ret = ''
    for s in sheets:
        ret += add(s) + '\n'
    return ret


def add(sheet):
    """
    Given the name of a spreadsheet, attempt to read 
    and process data, inserting it into RECORD_FILE
    """
    try:
        data = pd.read_excel(sheet)
    except:
        return f'FAIL: Could not read file {str(sheet)}'

    for col_name in [NAME] + COLS_TO_KEEP:
        if col_name not in data.columns:
            return f'FAIL: Did not find column {col_name}'

    # Remove multiple instances of names, keeping the first instance of each name
    new_data = data.groupby(NAME, as_index = False).first()
    new_data = new_data[[NAME] + COLS_TO_KEEP]
    new_data['file_name'] = sheet

    if RECORD_FILE in os.listdir():
        all_data = pd.read_csv(RECORD_FILE)
        if sheet in all_data.file_name.unique():
            return f"FAIL: {sheet} appears to have been added already"

        new_data = pd.concat([all_data, new_data], sort = False)

    new_data.to_csv(RECORD_FILE, index=False)

    return f"SUCCESS: Added {sheet} to records"


def create_excel():
    """
    Used by CLI ONLY
    Prompt for a filename, then output data into the file
    """
    output_file = input("Name of output file: ")
    if not output_file:
        return "FAIL: No output file given!"
    else:
        return create_excel_with_fname(output_file)


def create_excel_with_fname(fname):
    """ Given a filename, attempt to output data into that file """
    if RECORD_FILE not in os.listdir():
        return "FAIL: No data exists (hidden file '.data.csv' not found)"

    all_data = pd.read_csv(RECORD_FILE)
    all_data = all_data[[NAME] + COLS_TO_KEEP]

    final = {col: [] for col in [NAME, 'Counts'] + COLS_TO_KEEP}    
    name_count = Counter(all_data[NAME])
    for name in all_data[NAME].unique():
        final[NAME].append(name)
        final['Counts'].append(name_count[name])
        for col in COLS_TO_KEEP:
            good_data = list(set(all_data[col].loc[(~all_data[col].isna()) & (all_data[NAME] == name)]))
            final[col].append(good_data)
    final = pd.DataFrame(final)

    if fname[-5:] != ".xlsx":
        fname += ".xlsx"
    final.to_excel(fname)
    print(final)

    return f"SUCCESS: Created file {fname}, check terminal to peek at data"


def delete_event():
    """
    Used by CLI ONLY
    Prompt for filename, then remove any data in RECORD_FILE relating to the file
    """
    fname = None
    while not fname:
        fname = input( "Name of event to delete (enter 'list events' to view all events, " +
                            "'cancel' to return to main menu): ")
        if fname == 'list events':
            return list_events()
        elif fname == 'cancel':
            return "FAIL: Cancelled delete action"

    return delete_event_with_fname(fname)


def delete_event_with_fname(fname):
    """
    Given a filename, then remove any data in 
    RECORD_FILE relating to the file
    """
    if not os.path.exists(RECORD_FILE):
        return f"FAIL: {RECORD_FILE} does not exist"
    all_data = pd.read_csv(RECORD_FILE)

    all_data = all_data[all_data['file_name'] != fname]
    all_data.to_csv(RECORD_FILE, index=False)

    return f"SUCCESS: Deleted event '{fname}'"


def list_events():
    """
    List all events (filenames) which have 
    been added to RECORD_FILE
    """
    if RECORD_FILE not in os.listdir():
        return f"FAIL: No events to list because {RECORD_FILE} not found!"

    all_data = pd.read_csv(RECORD_FILE)
    events_list = f"\n".join([str(fname) for fname in all_data.file_name.unique()])

    return "Existing Events:\n" + events_list


def reset_records():
    """
    Used by CLI ONLY
    Prompt user for confirmation, then attempt to reset
    """
    if input("CONFIRM (enter 'yes'): ") == 'yes':
        return reset()


def reset():
    """
    Attempt to reset RECORD_FILE by deleting it
    """
    if RECORD_FILE in os.listdir():
        os.remove(RECORD_FILE)
        return "SUCCESS: Performed reset"
    else:
        return "FAIL: No reset necessary"


def quit_program():
    """
    Used by CLI ONLY
    Quit the program
    """
    sys.exit("Quitting")