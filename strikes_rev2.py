"""Strike management system using Google Sheets.

Prior to running, ensure that all Google Sheets information is correct.
Use after manually checking boxes noting who is present/absent.
All Google Sheets column numbers are hard-coded.

Author: Will Bradford
Version: 12/5/23

"""
import pygsheets
from colorama import Fore

# User input parameters (service account name and Google Sheets info)
serv_acc_filename = "INSERT SERVICE ACCOUNT FILE NAME HERE"
spreadsheet_name = "SPREADSHEET NAME"
attendance_worksheet_name = "ATTENDANCE WORKSHEET NAME"
strike_worksheet_name = "STRIKE WORKSHEET NAME"
# spreadsheet_name = "Python Test"
# attendance_worksheet_name = "Attendance Test"
# strike_worksheet_name = "Strike Test"

num_strikes_to_give = 1  # Number of strikes to give to each absentee

name_column_number = 1
attendance_column_number = 2
absentee_column_number = 9
excuse_column_letter = 'C'
present_column_letter = 'E'
late_column_letter = 'G'
absent_column_letter = 'I'
strikes_column_letter = 'B'

# Welcome message
print(f'{Fore.BLUE}Welcome to Astrinit, Will Bradford\'s attendance and strike management system.{Fore.RESET}')

# Create the Client
client = pygsheets.authorize(service_account_file=serv_acc_filename)

# Open attendance spreadsheet and worksheet
spreadsheet = client.open(spreadsheet_name)
attendance_worksheet = spreadsheet.worksheet("title", attendance_worksheet_name)
strike_worksheet = spreadsheet.worksheet("title", strike_worksheet_name)

# Get name and attendance columns, remove column titles, and standardize name format
name_column = attendance_worksheet.get_col(col=name_column_number, include_tailing_empty=False)
attendance_column = attendance_worksheet.get_col(col=attendance_column_number, include_tailing_empty=False)
all_names_standardized = [name.strip().lower() for name in name_column if name.strip().lower() != 'name']
attendance_standardized = [att for att in attendance_column if att != 'Present?']

name_attendance_dict = {name: attendance for name, attendance in zip(all_names_standardized, attendance_standardized)}

# Alphabetize capitalized attendees, create list to store capitalized absentees
attendees_capitalized = [name.title() for name in name_attendance_dict if name_attendance_dict[name] == 'TRUE']
absentees_capitalized = []

# Note final attendees and absentees
print(f'\nThe following people attended the meeting:\n{sorted(attendees_capitalized)}\n\n')
attendance_response = input('Begin noting attendance? Y/N:\n').lower().strip()

if attendance_response == 'y':

    print(f'\n\nAttendees and absentees are being noted. This could take awhile...\n\n')

    # Write attendees, absentees, and excused late people to Google Sheets. Account for title row
    attendee_walker = 2
    absentee_walker = 2
    late_walker = 2

    # Loop through everyone in club to note attendees and absentees
    for person in all_names_standardized:

        # Get excuse status. Account for indexing starting at 0 and deleting the name column title earlier
        person_excuse_status = attendance_worksheet.get_value(
            f'{excuse_column_letter}{all_names_standardized.index(person) + 2}').lower().strip()

        # If person is present or excused absent
        if (person.title() in attendees_capitalized) or person_excuse_status == 'excused absent':
            attendance_worksheet.update_value(f'{present_column_letter}{attendee_walker}', person.title())

            # Walk down column
            attendee_walker += 1

        # If person is excused late
        elif person_excuse_status == 'excused late':
            attendance_worksheet.update_value(f'{late_column_letter}{late_walker}', f'{person.title()}')

            # Walk down column
            late_walker += 1

        # If person is absent with no excuse
        else:
            attendance_worksheet.update_value(f'{absent_column_letter}{absentee_walker}', f'{person.title()}')

            # Add person to absentee list
            absentees_capitalized.append(person.title())

            # Walk down column
            absentee_walker += 1

        # Report progress
        print(f'{person.title()} accounted for.')

    # Report completion
    print(f'\n{Fore.LIGHTGREEN_EX}Successfully noted attendees and absentees.\n{Fore.RESET}')

# Get absentee column, remove column title, standardize name format, alphabetize list
absentee_column = attendance_worksheet.get_col(col=absentee_column_number, include_tailing_empty=False)
absentees_standardized = sorted([name.strip().lower() for name in absentee_column
                                 if name.strip().lower() != 'all unexcused absent'])

# Capitalize absentee names
absentees_capitalized = [name.title() for name in absentees_standardized]

# Start giving out strikes
strike_response = input(f'\nThe following people missed the {attendance_worksheet_name} meeting:\n'
                        f'{absentees_capitalized}\n\n'
                        f'Begin giving out strikes? Y/N:\n').lower().strip()

if strike_response == 'y':

    print(f'\n\nStrikes are being given out. This could take awhile...\n\n')

    # List to hold anyone at 3 strikes
    members_to_kick = []

    # Find absentee in roster, give strike
    for absentee in absentees_standardized:
        for member in all_names_standardized:
            if member == absentee:

                # Current row number. Account for indexing starting at 0 and deleting the name column title earlier
                row_num = all_names_standardized.index(member) + 2

                # Get num strikes
                member_num_strikes = strike_worksheet.get_value(f'{strikes_column_letter}{row_num}').strip()

                # Add strike(s) if at 1 or 2
                if member_num_strikes.isnumeric():
                    strike_worksheet.update_value(f'{strikes_column_letter}{row_num}',
                                                  f'{int(member_num_strikes) + num_strikes_to_give}')

                # Add first strike(s)
                else:
                    strike_worksheet.update_value(f'{strikes_column_letter}{row_num}', str(num_strikes_to_give))

                # Get new num strikes and check if any members are at 3 strikes. Compatible with any strike amount >= 3
                member_num_strikes = strike_worksheet.get_value(f'{strikes_column_letter}{row_num}').strip()
                if member_num_strikes >= '3':
                    members_to_kick.append(member.title())

                # Determine print color for strike number
                print_color = None
                if int(member_num_strikes) == 1:
                    print_color = Fore.LIGHTGREEN_EX
                elif int(member_num_strikes) == 2:
                    print_color = Fore.LIGHTYELLOW_EX
                else:
                    print_color = Fore.LIGHTRED_EX

                # Report progress
                print(f'Gave a strike to {member.title()}. {member.title()} now has '
                      f'{print_color}{member_num_strikes}{Fore.RESET} strike(s).')

                # Move to next absentee
                break

    # Report completion
    print(f'\n{Fore.LIGHTGREEN_EX}Successfully updated strikes.{Fore.RESET}\n')

    # Report if anyone at 3 strikes. Print in bold letters
    print(f'\033[1mThe following people have 3 strikes and should be kicked out:\n\033[0m{members_to_kick}')
