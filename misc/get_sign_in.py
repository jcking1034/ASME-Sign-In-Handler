#!/usr/bin/env python3
# Requires xlrd and openpyxl
import pandas as pd
import sys

NAME = 'Name'


if __name__ == '__main__':

    # Attempt to open specified spreadsheet
    read_from = input("Name of Excel Spreadsheet containing new data: ")
    try:
        data = pd.read_excel(read_from)
    except:
        raise Exception('Could not read file', read_from)

    # Remove multiple instances of names, keeping the first instance of each name
    new_data = data.groupby(NAME, as_index = False).first()

    # If given, try to open the target file
    prev_data_file = input( "Name of Excel Spreadsheet containing " +
                            "previous data (Press ENTER if no previous data): ")
    if prev_data_file:
        # Attempt to open specified spreadsheet
        try:
            prev_data = pd.read_excel(prev_data_file)
        except:
            raise Exception('Could not read file', prev_data_file)

        # Combine the basic information of both dataframes
        # NOTE: if data is updated in a new form, it won't be reflected
        # Basically, original data won't be overwritten
        all_info = pd.concat([new_data, prev_data], sort = False)
        final = all_info.groupby(NAME, as_index = False).first()

    # Otherwise, nothing to combine
    else:
        final = new_data
    
    # Sort in alphabetical order by name
    final = final.sort_values(by=[NAME])

    # display results and save to a file
    print("\nResulting Data:\n", final, "\n")
    output_file = input("Name of output file: ")
    if output_file:
        final.to_excel(output_file)
    else:
        raise Exception("No output file given!")

    print("Successfully executed processing!")