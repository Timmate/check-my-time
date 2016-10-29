#! python3
#
# NAME          : check_my_time.py
#
# VERSION       : 2.0
#
# DESCRIPTION   : A simple time and attendance system.
#
# AUTHOR        : Tim Kornev (@Timmate on GitHub)
#
# CREATED DATE  : 13th of July, 2016
#
# LAST MODIFIED : 29th of October, 2016
#


import sys    # is used to exit script in case of fatal error
import os
import ast    # is used to parse data from text files
import datetime
import pprint    # pretty prints data to text files


# Constants.
# ==========

TODAY = datetime.date.today()    # is used by all functions in the script

# `%-d` is non-zero-padded day's number. Could be changed to `%d` for zero-padded day's number
# NOTE: if changed to zero-padded option, also change `ZERO_PADDED_FILENAMES` in `report_creator.py`
# from `False` to `True`. Otherwise, Report Creator will not be able to 'see' some files.
TODAY_FILENAME = TODAY.strftime('%-d.txt')

SAVE_PROMPT = True    # change to `False` to cancel manual input data confirmation

DEFAULT_START_TIME_HOUR = 9    # note that it's 24-hour time system, not 12-hour with AM's and PM's
DEFAULT_START_TIME_MINUTE = 0    # note that number must be non-zero padded. So use `0` instead of `00` or `7` instead of `07`

# Text files are stored in /.../Work Attendance Files/<Year>/<Month's number> — <Month>/
WORKING_DIR = 'Work Attendance Files'
YEAR_STRF = TODAY.strftime('%Y')
YEAR_DIR = os.path.join(WORKING_DIR, YEAR_STRF)

# `%-m` is month's number. This helps file managers sort dirs. Do NOT change
# this value in parentheses as it will cause an error to Report Creator.
MONTH_STRF = TODAY.strftime('%-m — %B')

MONTH_DIR = os.path.join(YEAR_DIR, MONTH_STRF)
PATH_TO_FILENAME = os.path.join(MONTH_DIR, TODAY_FILENAME)    # is used by `load_data()`, `write_to_file()`

# Functions.
# ==========

def main():
    """The core function."""

    # This global is used by almost all functions.
    global args

    display_menu()
    load_data()

    try:
        # Ask for input until Ctrl-C is pressed.
        while True:
            print()

            input_data = input("Enter name and time: ")
            args = input_data.split()

            if input_data.strip().upper() == 'ALL':
                display_present_workers()

            elif input_data.strip().upper() == 'MENU':
                display_menu()

            else:
                if validate_data():    # if the input data is correct
                    if name not in data:    # if a name was entered for the first time a day
                        clock_in()
                    else:    # if a name was entered for the second/third time a day
                        clock_out()

    except KeyboardInterrupt:    # handle Ctrl-C exception
        write_to_file()




def load_data():
    """
    Loads data from a text file if it exists. Otherwise creates a new dictionary
    that will store data.
    """

    # These globals are used by almost all functions.
    global data, day_start_dt


    if os.path.exists(PATH_TO_FILENAME):
        # Open file for today and load data.
        with open(PATH_TO_FILENAME) as f:
            print()
            print('Loading data from file {} ...'.format(TODAY_FILENAME))

            data = ast.literal_eval(f.read())

        if 'day_start' in data:
            if 'day_start_dt' in data['day_start']:
                day_start_dt = eval(data['day_start']['day_start_dt'])

        else:
            # Fatal error. No calculations can be made, so exit.
            print('Error: no `day_start_dt` value found.')
            sys.exit(0)

    else:
        # Create a new dictionary that will store data.
        data = {}

        # Allow the user to set new day's start time.
        hour, minute = DEFAULT_START_TIME_HOUR, DEFAULT_START_TIME_MINUTE

        choice = None

        while choice not in ('OK', 'SET'):
            print('DATE: {}.'.format(TODAY.strftime('%d %b %Y')))
            print('DEFAULT START TIME: {}:{}.'.format(hour, minute))

            choice = input("Type 'OK' to continue with default start time or "
                                      "'SET' to set new start time: ")
            choice = choice.upper()

        if choice == 'SET':
            # Set new start time.
            while True:
                try:
                    print()

                    hour = int(input('Enter new start hour (0..23): '))
                    minute = int(input('Enter new start minute (0..59): '))
                    datetime.time(hour, minute)    # use datetime.time to validate entered time

                    break

                except ValueError as exc:    # if datetime.time validation fails
                    print('Error: ' + str(exc))

                    continue

        day_start_dt = datetime.datetime(TODAY.year, TODAY.month, TODAY.day, hour, minute)

        # Add new start time to data so that it could be used when loaded later.
        data['day_start'] = {'day_start_dt': pprint.pformat(day_start_dt),    # or day_start_time_dt, day_start_hour
                             'day_start_hour': hour,
                             'day_start_minute': minute
                             }


    # Display day's start time.
    day_start_strf = day_start_dt.strftime('=== %d %b %Y ===  %H:%M ===')
    day_start_strf = day_start_strf.center(75)     # center alignment

    print()
    print("DAY'S START: {}".format(day_start_strf))




def validate_data():
    """Validate input data."""

    # These globals are used by these functions: `clock_in()`, `clock_out()`
    global name, time_argument, hour, minute

    try:
        if len(args) == 1:    # only name was entered
            name = args[0]
            assert name.isalpha(), 'name must contain only letters.'

            time_argument = False

            return True

        elif len(args) > 1:    # name and time were entered or just full name was entered
            time_argument = True    # assume that time argument was entered

            # Validate time argument.
            try:
                # If this fails, `ValueError` is raised, saying that
                # "there are not enough values to unpack (expected 2, got 1)"
                # That means that no time argument but full name was entered.
                hour, minute = args[-1].split(':')

                # If this fails, `ValueError` is raised, saying that
                # "invalid literal for int() with base 10"
                hour, minute = int(hour), int(minute)

                # If this fails, `ValueError` is raised, saying that
                # "hour/minute must be in ..."
                datetime.time(hour, minute)

            except ValueError as err:
                if 'unpack' in str(err):
                    name = ' '.join(args)    # full name, to be exact
                    time_argument = False

                else:
                    print('Error: ' + str(err))

                    return False

            else:    # is executed only if the statements in `try` block do not raise an exception
                name = ' '.join(args[:-1])    # name/full name without the time argument

            # Ensure that the name contains only letters.
            for part_of_name in name.split():
                assert part_of_name.isalpha(), 'name must contain only letters.'

            return True

        else:
            # Throw `AssetionError` instead of `Exception` to avoid hiding bugs.
            raise AssertionError('whitespace is not a name.')     # if whitespace was entered

    except AssertionError as err:    # handle `AssertionError` from names validation
        print('Error: ' + str(err))

        return False




def clock_in():
    """Records time a person clocked in."""

    global data

    # Find time the person clocked in.
    if len(args) == 1 or not time_argument:
        # Use current time.
        clock_in_dt = datetime.datetime.now()
        clock_in_strf = clock_in_dt.strftime('%H:%M')

    elif len(args) > 1 and time_argument:
        # Use time argument from input.
        year, month, day = day_start_dt.year, day_start_dt.month, day_start_dt.day
        clock_in_dt = datetime.datetime(year, month, day, hour, minute)
        clock_in_strf = str(hour) + ':' + str(minute)    # this value comes handy when inspecting text files

    clock_in_early = False    # assume that the person did not come before day's start time

    if clock_in_dt < day_start_dt:    # if the person did come before before day's start time
        # Ask for the user's confirmation.
        choice = None

        while choice not in ('y', 'n'):
            choice = input('Are you sure "{}" clocked in before day\'s start time and '
                           'was not late? [y/n]: '.format(name))
            choice = choice.lower()

        if choice == 'y':
            clock_in_early = True
        else:
            return None    # brings back to infinite `while` loop

    if clock_in_early:
        # Calculate time the person was early.
        early_time_td = day_start_dt - clock_in_dt
        seconds = early_time_td.seconds

        early_time_hour = seconds // 3600

        # Display a message to the user.
        if early_time_hour > 0:
            early_time_minute = seconds % 3600 // 60

            print('"{}" clocked in in at {} and was {} hour(s), {} minute(s) early.'\
                  .format(name, clock_in_strf, early_time_hour, early_time_minute))

        else:
            early_time_minute = seconds // 60

            print('"{}" clocked in in at {} and was {} minute(s) early.'\
                  .format(name, clock_in_strf, early_time_minute))

    else:
        # Calculate time the person was late for.
        late_time_td = clock_in_dt - day_start_dt
        seconds = late_time_td.seconds

        late_time_hour = seconds // 3600

        # Display a message to the user.
        if late_time_hour > 0:
            late_time_minute = seconds % 3600 // 60
            print('"{}" clocked in at {} and was late for {} hour(s), {} minute(s).'\
                  .format(name, clock_in_strf, late_time_hour, late_time_minute))
        else:
            late_time_minute = seconds // 60
            print('"{}" clocked in at {} and was late for {} minute(s).'\
                  .format(name, clock_in_strf, late_time_minute))

    # Write data to the dictionary.
    if SAVE_PROMPT:
        # Ask for the user's confirmation.
        choice = None

        while choice not in ('y', 'n'):
            choice = input('Do you want to proceed? [y/n]: ')
            choice = choice.lower()

        if choice == 'n':
            return None    # brings back to asking infinite `while` loop

        elif choice == 'y':
            data[name] = {# Despite we do not really need this value in calculations, it
                          # comes handy when inspecting text files.
                          'clock_in_early': clock_in_early,
                          
                          'clock_in_dt': pprint.pformat(clock_in_dt),

                          # Despite we do not really need this value in calculations, it
                          # comes handy when inspecting text files.
                          'clock_in_strf': clock_in_strf
                          }

            if clock_in_early:
                # Write 'early' data only.
                data[name]['early_time_hour'] = early_time_hour
                data[name]['early_time_minute'] =  early_time_minute

            else:
                # Write 'late' data only.
                data[name]['late_time_hour'] = late_time_hour
                data[name]['late_time_minute'] =  late_time_minute

    else:    # continue without confirmation
        data[name] = {'clock_in_early': clock_in_early,
                      'clock_in_dt': pprint.pformat(clock_in_dt),
                      'clock_in_strf': clock_in_strf
                      }

        if clock_in_early:
            # Write 'early' data only.
            data[name]['early_time_hour'] = early_time_hour
            data[name]['early_time_minute'] =  early_time_minute

        else:
            # Write 'late' data only.
            data[name]['late_time_hour'] = late_time_hour
            data[name]['late_time_minute'] =  late_time_minute




def clock_out():
    """Records time a person clocked out."""

    global data

    if 'work_time_hour' not in data[name]:    # if the person has NOT already clocked out
        # Find time the person clocked out.
        if len(args) == 1:
            # Use current time.
            clock_out_dt = datetime.datetime.now()
            clock_out_strf = clock_out_dt.strftime('%H:%M')    # this value comes handy when inspecting text files

        elif len(args) > 1:
            # Use time argument from input.
            year, month, day = day_start_dt.year, day_start_dt.month, day_start_dt.day
            clock_out_dt = datetime.datetime(year, month, day, hour, minute)
            clock_out_strf = str(hour) + ':' + str(minute)    # this value comes handy when inspecting text files

        # Find time the person clocked in.
        clock_in_dt = eval(data[name]['clock_in_dt'])

        # Prevent incorrect input.
        if clock_out_dt < day_start_dt or clock_out_dt < clock_in_dt:
            print('"{}" could not clock out at that time.'.format(name))
            return None    # brings back to infinite `while` loop

        # Calculate time of working.
        work_time_td = clock_out_dt - clock_in_dt
        seconds = work_time_td.seconds
        work_time_hour = seconds // 3600

        # Display a message to the user.
        if work_time_hour > 0:
            work_time_minute = seconds % 3600 // 60
            print('"{}" clocked out at {} and worked for {} hour(s), {} '
                  'minute(s).'.format(name, clock_out_strf, work_time_hour,
                                      work_time_minute))

        else:
            work_time_minute = seconds // 60
            print('"{}" clocked out at {} and worked for {} '
                  'minute(s).'.format(name, clock_out_strf, work_time_minute))

        if SAVE_PROMPT:
            # Ask for confirmation.
            choice = None
            while not choice in ('y', 'n'):
                choice = input('Do you want to proceed? [y/n]: ')
                choice = choice.lower()

            if choice == 'n':
                return None    # brings back to infinite `while` loop
            elif choice == 'y':
                data[name].pop('clock_in_dt')  # there is no need to store it anymore
                data[name]['clock_out_strf'] = clock_out_strf
                data[name]['work_time_hour'] = work_time_hour
                data[name]['work_time_minute'] = work_time_minute

        else:    # continue without confirmation
            data[name].pop('clock_in_dt')  # there is no need to store it anymore
            data[name]['clock_out_strf'] = clock_out_strf
            data[name]['work_time_hour'] = work_time_hour
            data[name]['work_time_minute'] = work_time_minute

    elif 'work_time_hour' in data[name]:
        # If the person has already clocked out and left workplace.
        print('{} has already left workplace.'.format(name))




def display_present_workers():
    """Displays present workers."""

    # NOTE: `clock_in_dt` is popped after the person leaves workplace so it helps to
    # detect present workers. But do not use `not work_time_hour` here as it displays
    #`day_start` dict too.
    present_workers = [name for name in data if 'clock_in_dt' in data[name]]

    print()
    print('NOW ON WORKPLACE:')

    if present_workers:
        for name in present_workers:
            print('\t' + name)
    else:
        print('\t NOBODY')




def display_menu():
    """Displays menu."""

    menu = ('\t\t' + "+——————————————————————————————————————————————————————————————————————————————————+" + '\n'
            '\t\t' + "| • Press ENTER to close this menu.                                                |" + '\n'
            '\t\t' + "| • Enter a person's name the first time a day to record time he/she clocked in.   |" + '\n'
            '\t\t' + "|   (e.g. >>> Name (Full Name) (hh:mm))                                            |" + '\n'
            '\t\t' + "| • Enter a person's name the second time a day to record time he/she clocked out. |" + '\n'
            '\t\t' + "|   (e.g. >>> Name (Full Name) (hh:mm))                                            |" + '\n'
            '\t\t' + "| • Type ALL to see a list of all workers present on workplace.                    |" + '\n'
            '\t\t' + "| • Type MENU to display this menu.                                                |" + '\n'
            '\t\t' + "| • Press Ctrl-C to write data to a file and exit.                                 |" + '\n'
            '\t\t' + "+——————————————————————————————————————————————————————————————————————————————————+"
            )

    print()
    print(menu)
    input()




def write_to_file():
    """Writes data to a file."""

    # Create a month dir. `MONTH_DIR` constant is defined on top of the script.
    os.makedirs(MONTH_DIR, exist_ok=True)

    with open(PATH_TO_FILENAME, 'w') as f:
        f.write(pprint.pformat(data))

        # Display the filename and path to it.
        print()
        print('Saved as "{}" to "{}"'.format(TODAY_FILENAME, MONTH_DIR))
        print()




if __name__ == '__main__':
    main()
