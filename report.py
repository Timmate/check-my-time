#! python3
#
# NAME          : report.py
#
# DESCRIPTION   : Creates a week or month work attendance report based on
#                 data from the files created with work_attendance.py and
#                 stored in \wa_files directory. Writes report data to an
#                 Excel spreadsheet and saves it to \reports directory.
#
# AUTHOR        : Tim Kornev (Timmate profile on GitHub)
#
# CREATED DATE  : 19th of July, 2016
#
# TODO:
#       - Design output to the Excel spreadsheet (i.e. cell's width)
#


import sys
import os
import datetime

from functions import collect_data, sum_data
from functions import calc_time_of_working, format_late_time
from functions import write_to_xl


# Change CWD here.
BASE_DIR = r'C:\Users\<User>\Desktop'
WORK_FOLDER = 'wa_files'
PATH = os.path.join(BASE_DIR, WORK_FOLDER)
os.chdir(PATH)
sys.path.append(PATH)   # this is needed for importing month/week modules

NO_FILES = False        # if no files found, display a message and do not create a report

# Display the CWD.
print()
print('CWD: {}'.format(os.getcwd()))
print()

# Ask the user for choice.
month_or_week = None

while not month_or_week in ('MONTH', 'WEEK'):
    month_or_week = input('Enter MONTH or WEEK to create a review.')
    
while True:
    try:
        date_as_string = input('Enter the date: dd/mm/YYYY')
        date_as_datetime_obj = datetime.datetime.strptime(date_as_string, '%d/%m/%Y')
        break   # If date format is correct, break the 'while' loop.
    
    except ValueError:
        print('ERROR: Invalid date value.')
        continue


# MONTH.
# ======

if month_or_week == 'MONTH':
    # Find first and last days of the month.
    year, month = date_as_datetime_obj.year, date_as_datetime_obj.month
    start = datetime.datetime(year, month, 1)
    end = datetime.datetime(year, month + 1, 1) - datetime.timedelta(days=1)

    # Read in data from all files and write it to one file.
    collect_data('month', start, end)

    # Import all persons dicts from a month_data file in order to use them in calculations.
    print()
    print('Reading month data from month_data_file...')
    print()
    month_data = {}    # Contain sum of all month data in this dictionary.
    # NOTE: For those days that do not have a 'wa_...' file empty dicts will
    # be created.
    from month_data_file import persons_1, persons_2, persons_3, persons_4, persons_5
    from month_data_file import persons_6, persons_7, persons_8, persons_9, persons_10
    from month_data_file import persons_11, persons_12, persons_13, persons_14, persons_15
    from month_data_file import persons_16, persons_17, persons_18, persons_19, persons_20
    from month_data_file import persons_21, persons_22, persons_23, persons_24, persons_25
    from month_data_file import persons_26, persons_27, persons_28, persons_29, persons_30, persons_31

    # The actual calculation process comes here.
    persons_lst = [persons_1, persons_2, persons_3, persons_4, persons_5,
                   persons_6, persons_7, persons_8, persons_9, persons_10,
                   persons_11, persons_12, persons_13, persons_14, persons_15,
                   persons_16, persons_17, persons_18, persons_19, persons_20,
                   persons_21, persons_22, persons_23, persons_24, persons_25,
                   persons_26, persons_27, persons_28, persons_29, persons_30, persons_31
                   ]

    for persons_dict in persons_lst:
        sum_data(month_data, persons_dict, start)
    calc_time_of_working(month_data, start)
    format_late_time(month_data)    

    if not month_data:
        NO_FILES = True
        print()
        print('NO FILES FOUND.')
    else:
        # Write result to an Excel spreadsheet.
        write_to_xl(month_data, 'month', start, end)
        
    # Delete month_data_file. (You can comment this statement if
    # you want to look at the data in this file.)
    os.remove('month_data_file.py')    

  
# WEEK.
# =====

elif month_or_week == 'WEEK':
    weekday = date_as_datetime_obj.weekday()
    start = date_as_datetime_obj - datetime.timedelta(days=weekday)
    start = datetime.datetime(start.year,
                              start.month,
                              start.day)
    end = start + datetime.timedelta(days=6)

    # Read in data from all files and write it to one file.
    collect_data('week', start, end)

    # Import all persons dicts from a week_data file in order to use them in calculations.
    print()
    print('Reading week data from week_data_file...')
    week_data = {}   # Contain sum of all week data in this dictionary.
    from week_data_file import persons_1, persons_2, persons_3, persons_4
    from week_data_file import persons_5, persons_6, persons_7

    # The actual calculation process comes here.
    persons_lst = [persons_1, persons_2, persons_3, persons_4, persons_5,
                   persons_6, persons_7
                   ]

    for persons_dict in persons_lst:
        sum_data(week_data, persons_dict, start)
    calc_time_of_working(week_data, start)
    format_late_time(week_data)

    if not week_data:
        NO_FILES = True
        print()
        print('NO FILES FOUND.')
    else:
        # Write result to an Excel spreadsheet.
        write_to_xl(week_data, 'week', start, end)
        
    # Delete week_data_file. (You can comment this statement if
    # you want to look at the data in this file.)
    os.remove('week_data_file.py')


# Create the report.
# ==================

if not NO_FILES:
    # Display some information.
    start_strf = start.strftime('%d_%b_%Y')
    end_strf = end.strftime('%d_%b_%Y')

    print()
    print('FROM: {}'.format(start_strf))
    print('TO: {}'.format(end_strf))
    print()
    print('Creating {} report...'.format(month_or_week.lower()))
    print()
    print('Done.')
    print()
else:
    print('CANNOT CREATE A {} REPORT.'.format(month_or_week))
    print()
