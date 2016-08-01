#! python3
#
# NAME         : work_attendance.py
#
# VERSION      : 1.0
#
# DESCRIPTION  : A simple timesheet app that records when the user types
#                a person's name and uses the current or specified time
#                to clock them in or out.
#
# AUTHOR       : Tim Kornev (Timmate profile on GitHub)
#
# CREATED DATE : 13th of July, 2016
#
# TODO:
#       - Create a directory for wa_files (i.e. June 2016) and contain
#         all files for that month and year there. It will add some
#         convenience.
#


import os
import sys
import datetime
import pprint


# Write all data to a new file or import data from the
# existing file in /wa_files directory.
BASE_DIR = r'C:\Users\Тима\Desktop'
SAVE_FOLDER = 'wa_files'

os.chdir(BASE_DIR)
os.makedirs(SAVE_FOLDER, exist_ok=True)  # Make sure the folder exists.

PATH = os.path.join(BASE_DIR, SAVE_FOLDER)
os.chdir(PATH)      # Script could be run everywhere, but
                    # save folder is on the desktop.
sys.path.append(PATH)    # Path should be added so that program is able
                         # to import the module with 'persons' dict.


# Display CWD and the folder where all files will be saved.
print()
print('CWD: {}'.format(BASE_DIR))
print('SAVE FOLDER: {}.'.format(SAVE_FOLDER))



# Display program instructions.
print()
print(" ----------------------------------------------------------------")
print("| • Press ENTER to begin.                                        |")
print("| • Type the person's name the first time to clock him/her in.   |")
print("|   (ex. >>> Name (Full name ) (hh:mm))                          |")
print("| • Type the person's name the second time to clock him/her out. |")
print("|   (ex. >>> Name (Ful name) (hh:mm))                            |")
print("| • Type ALL to see the list of all workers on the workplace.    |")
print("| • Type DEL to delete datetime objects from all workers.        |")  # del clock_in_datetime in case
print("| • Press Ctrl-C to exit.                                        |")  # the worker won't come to the
print(" ---------------------------------------------------------------- ")  # workplace that day.
input()   # press ENTER to begin.



# Set date and time for beginning of the working day.
NOW = datetime.datetime.now()
START_OF_DAY = datetime.datetime(NOW.year, NOW.month, NOW.day, 9)  # Start day at 9 AM as default.

start_of_day_strf = START_OF_DAY.strftime('%d_%b_%Y')
filename = 'wa_{}.py'.format(start_of_day_strf)   # use a suitable string format for the filename.


if filename in os.listdir('.'):
    # Open file for today and load all data.
    print('Loaded data from {}.'.format(filename))
    # Since we need to be able to get data from the file and use it,
    # data should be written to a .py file that could be imported as
    # a Python module.
    # NOTE: Module must be imported using a fixed name, not a variable.
    # This name of the module could be changed automatically with
    # change_date.py (see it's docs for details). change_date.py
    # must be run in the beginning of the every day.
    # Otherwise, the app will work only with one day (now it is set to 16/06/2016).
    from wa_16_Jul_2016 import persons
else:
    # Create a new dictionary that will contain all data for today.
    persons = {}  # store persons' names and their work attendance.

    # Allow the user to set start hour (only if the app was launched
    # first time for today).
    choice = None
    while True:
        print()
        print('START DATE: {}.'.format(start_of_day_strf))
        print('START TIME: 9 AM.')
        choice = input("Type 'OK' or enter new start hour (9 - 12)" + \
                       " and minute: (hh:mm)")

        if choice.upper() == 'OK':
            # continue with original time.
            break

        elif len(choice) == 5 and \
             choice[:2].isdigit() and choice[3:].isdigit() and \
             8 < int(choice[:2]) < 13 and 0 <= int(choice[3:]) < 60:
            # continue with new time set.
            new_hour = int(choice[:2])
            new_minute = int(choice[3:])
            START_OF_DAY = datetime.datetime(NOW.year,
                                               NOW.month,
                                               NOW.day,
                                               new_hour,
                                               new_minute)
            break


# Display the start date and time.
started_string = START_OF_DAY.strftime('%d/%b/%Y --- %H:%M')
print()
print('STARTED: {}'.format(started_string))
print()


# Ask user for input.
try:
    while True:
        # Check whether the input was correct and ask the
        # user for another one if not.
        name = input('Type the person\'s name and time (optional). ')
        args = name.split(' ')
        if len(args) > 1:
            # i.e. Alice Margaret Mitchell 12:08 or Margaret 13:22
            name = ' '.join(args[:-1])  # Full name, to be exact.
            try:
                hours = int(args[-1][:2])
                minutes = int(args[-1][3:])
            except ValueError:
                print('ERROR: Ivalid time input.')
                continue

        try:
            if len(args) == 1:
                # If only one word for name was entered.
                assert name.isalpha(), 'Name must contain only letters.'

            elif len(args) > 1:
                # If full name was entered and/or time argument.

                # Make sure (full) name contains only letters.
                for part_of_full_name in args[:-1]:
                    assert part_of_full_name.isalpha(), 'Name must contain only letters.'

                # Make sure the time argument is valid according to working day start time.
                assert 0 <= minutes < 60, 'Number of minutes must be in range(0, 60)'
                assert START_OF_DAY.hour <= hours < 24, 'Number of hours' + \
                                                          'must be in range({}, 24)'.format(START_OF_DAY.hour)
                if START_OF_DAY.hour == hours:
                    # i.e. Day began at 9.30 but a worker clocked in at 9.29
                    # NOTE: This is not needed now as start hour value is integer.
                    assert START_OF_DAY.minute < minutes, 'The worker could not clock in' + \
                                                            'or clock out at that time.'
        except AssertionError as err:
            print()
            print('ERROR: ' + str(err))
            print()
            continue



        if not name in persons.keys():
            if name == 'ALL':
                # Show who came to the workplace.
                who_came = [name for name in persons.keys() if (not persons[name]['clock_out'])]
                print()
                print('NOW ON THE WORKPLACE:')
                if who_came:
                    for person in who_came:
                        print(person)
                else:
                    print('NOBODY')

                print()
                continue

            elif name == 'DEL':
                # Delete datetime objects from all 'persons' dictionary values.
                for name in persons.keys():
                    if not persons[name]['clock_out']:
                        print()
                        print('DELETED FROM: {}'.format(name))
                        print()
                        persons[name].pop('clock_in_datetime')
                        continue

            else:
                # The person have just come to workplace. (if a real name was typed.)
                # Investigate how late he or she was and add data to the 'persons' dict.
                if len(args) > 1:
                    # If time argument also was entered.
                    clock_in_datetime = datetime.datetime(NOW.year,
                                                          NOW.month,
                                                          NOW.day,
                                                          hours,
                                                          minutes)
                    clock_in_strf = args[-1]
                else:
                    # If only name was typed.
                    clock_in_datetime = datetime.datetime.now()
                    clock_in_strf = clock_in_datetime.strftime('%H:%M')

                late_for_delta = clock_in_datetime - START_OF_DAY
                late_for_minutes = late_for_delta // 60

                print('{} came at {} and was late for {} minute(s). '.format(name,
                                                                             clock_in_strf,
                                                                             late_for_minutes))

                persons[name] = {'clock_in': clock_in_strf,
                'clock_in_datetime': clock_in_datetime,
                'late_for': '{} min(s)'.format(late_for_minutes),
                'clock_out': None,   # This will be overwritten if the person's name is typed twice.
                'worked_for': None}  # This will be overwritten as well.



        elif name in persons.keys():
            # The person have just left the workplace.
            # Investigate when the person ended their working day and how much time
            # they worked for. Add that data to the 'persons' dict.
            try:
                clock_in_datetime = persons[name]['clock_in_datetime']

                if len(args) > 1:
                    clock_out_datetime = datetime.datetime(NOW.year,
                                                           NOW.month,
                                                           NOW.day,
                                                           hours,
                                                           minutes)
                    clock_out_strf = args[-1]
                else:
                    clock_out_datetime = datetime.datetime.now()
                    clock_out_strf = clock_out_datetime.strftime('%H:%M')

                worked_for_delta = clock_out_datetime - clock_in_datetime
                worked_for_hours = worked_for_delta.seconds // 3600
                worked_for_minutes = (worked_for_delta.seconds % 3600) // 60
                worked_for_hours_and_mins = '{}:{}'.format(worked_for_hours,
                                                           worked_for_minutes)

                print('{} left work at {} and worked for {}.'.format(name,
                                                                     clock_out_strf,
                                                                     worked_for_hours_and_mins))

                persons[name].pop('clock_in_datetime')  # There is no need to store it anymore.
                persons[name]['clock_out'] = clock_out_strf
                persons[name]['worked_for'] = worked_for_hours_and_mins

            except KeyError:
                # datetime object was deleted manually with DEL by the user.
                if not persons[name]['clock_out']:
                    print('Datetime Object for {} was deleted.'.format(name))
                    continue

                # If the person's name was typed three times during one day.
                print('{} has already left the workplace.'.format(name))
                continue



except KeyboardInterrupt:
    # Write all data to a file.
    print()
    print('Writing data to {} file...'.format(filename))
    print()
    write_file = open(filename, 'w')
    write_file.write('import datetime\n')    # This is needed for exporting persons
                                             # with clock_in_datetime set to Datetime Object.
                                             # Otherwise no datetime module will be found and
                                             # an error will raise.
    write_file.write('persons = {\n\n ' + pprint.pformat(persons)[1:])
    write_file.close()
    print('Done.')
    print()
