#!/usr/bin/env python3
import utils

if __name__ == '__main__':
    command_list = {
        "help":                 utils.display_help,
        "add":                  utils.add_spreadsheet,
        "add from directory":   utils.add_from_directory,
        "output":               utils.create_excel,
        "delete":               utils.delete_event,
        "list events":          utils.list_events,
        "reset":                utils.reset_records,
        "quit":                 utils.quit_program
    }

    while True:
        try:
            command = input("\nEnter Command ('help' for help): ")
            if command in command_list:
                print(f"\n*****\n{command}")
                print(command_list[command]())
            else:
                print(f"UNRECOGNIZED COMMAND: {command}")
        except KeyboardInterrupt:
            print("\nEnter 'quit' to quit")