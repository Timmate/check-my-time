#! python3
#
# NAME         : functions.py
#
# DESCRIPTION  : Contains functions that are used by report.py script.
#
# AUTHOR       : Tim Kornev (Timmate profile on GitHub)
#
# CREATED DATE : 17th of July, 2016
#


import os
import sys
import datetime

import openpyxl
from openpyxl.styles import Font


BASE_DIR = r'C:\Users\Тима\Desktop'
WORK_FOLDER = 'wa_files'
SAVE_FOLDER = 'reports'     # Contain all reports in \reports.
PATH = os.path.join(BASE_DIR, WORK_FOLDER)    # Path to files.
PATH_TO_REPORTS = os.path.join(BASE_DIR, SAVE_FOLDER)

os.makedirs(PATH_TO_REPORTS, exist_ok=True)  # Make sure the \reports folder exists.
os.chdir(PATH)
sys.path.append(PATH)


def collect_data(period, start, end):
    '''
    Gathers data from all files for the given period and
    writes it to one file.
    '''
    day = 0   # Will be used to add to each persons dict a number of
              # week/month day, starting from 0.

    if period == 'month':
        write_file = open('month_data_file.py', 'w')   # Data will be written to this file.

    elif period == 'week':
        write_file = open('week_data_file.py', 'w')    # The same thing.

    else:
        raise Exception('ERROR: Entered invalid period value.')

    start_strf = start.strftime('%d_%b_%Y')
    end_strf = end.strftime('%d_%b_%Y')
    print()
    print('Opening write_file...')

    current_date = None

    while current_date != end:
        current_date = start + datetime.timedelta(days=day)
        current_date_strf = current_date.strftime('%d_%b_%Y')
        filename = 'wa_{}.py'.format(current_date_strf)

        if os.path.exists(os.path.join(PATH, filename)):
            print('Reading in {} file...'.format(filename))
            read_file = open(filename)
            contents = read_file.read()
            read_file.close()
            # Erase these statements from the file so that week/month_data_file
            # contains only persons_ dictionaries.
            contents = contents.replace('persons', 'persons_' + str(day + 1))
            contents = contents.replace('import datetime', '')
            write_file.write(contents)
        else:
            # Display missing files' names.
            print('Missing file: {}...'.format(filename))
            print()
            write_file.write("\npersons_" + str(day + 1) + " = {} \n")

        day += 1


    if period == 'week' and day < 7:   # If any files for that week
                                       # do not exist, write empty dicts
                                       # to the file.
        for number in range(day + 1, 8):
            write_file.write("\npersons_" + str(number) + " = {} \n")

    elif period == 'month' and day < 31:
        for number in range(day + 1, 32):
            write_file.write("\npersons_" + str(number) + " = {} \n")

    write_file.close()
    print('Closing write_file...')


def sum_data(container, persons_dict, start):
    '''
    Adds data from persons_dict to the container.
    '''
    for person in persons_dict.keys():
        container.setdefault(person, {'late_for_mins_summary': 0,
                                      'worked_until_date': start}) # 'worked_until_date' is a datetime
                                                                   # object and will be used for computing
                                                                   # actual time the person worked for.

        for key, value in persons_dict[person].items():
            if key == 'late_for':
                # Find minutes value from a string like '345 min(s)' and add it to sum.
                minutes = int(value.split(' ')[0])
                container[person]['late_for_mins_summary'] += minutes

            elif key == 'worked_for':
                # Find hours and minutes value from a string like '12:05' and add them to sum.
                hours, minutes = int(value.split(':')[0]), int(value.split(':')[1])
                delta = datetime.timedelta(hours=hours, minutes=minutes)
                container[person]['worked_until_date'] += delta

            else:
                # We do not need 'clock_in' and 'clock_out' keys from dicts.
                continue

    return container


def calc_time_of_working(container, start):
    '''
    Calculates amount of time the person has worked for.
    '''
    for person in container.keys():
        for key in list(container[person]):  # list() should be used here as
                                             # dict's size changes.
            if key == 'worked_until_date':
                # Extract datetime of the working period start date
                # from the date the person worked until.
                time_worked_for = container[person][key] - start
                hours = int(time_worked_for.total_seconds() // 3600)
                minutes = int((time_worked_for.total_seconds() % 3600) // 60)
                hours_and_mins_str = '{}:{}'.format(str(hours), str(minutes))
                container[person]['worked_for_summary'] = hours_and_mins_str
                # There is no need to store this key anymore.
                container[person].pop(key)

    return container


def format_late_time(container):
    '''
    Format late time in minutes to hours and minutes.
    '''
    for person in container.keys():   
        for key in list(container[person]):
            if key == 'late_for_mins_summary':
                minutes = container[person][key]
                hours = 0
                if minutes > 60:
                    hours = minutes // 60
                    minutes = minutes % 60
                hours_and_mins = '{}:{}'.format(hours, minutes)
                container[person]['late_for_summary'] = hours_and_mins

    return container


def write_to_xl(container, period, start, end):
    '''
    Writes all data from the container to an Excel spreadsheet.
    '''

    wb = openpyxl.Workbook()
    sheet = wb.active

    font_obj = Font(sz=13, bold=True)  # Use this font for labels.

    sheet['A1'] = 'Name'
    sheet['B1'] = 'Late for'
    sheet['C1'] = 'Worked for'
    sheet['A1'].font = font_obj
    sheet['B1'].font = font_obj
    sheet['C1'].font = font_obj

    row = 2

    for person in container.keys():
        sheet['A' + str(row)] = person
        sheet['B' + str(row)] = container[person]['late_for_summary']
        sheet['C' + str(row)] = container[person]['worked_for_summary']
        row += 1

    if period == 'month':
        month_and_year = start.strftime('%b_%Y')
        filename = '{} wa_report.xlsx'.format(month_and_year)       # The format is: <month>_<year>
    elif period == 'week':
        week_beginning = start.strftime('%d.%m')
        week_end = end.strftime('%d.%m')
        filename = '{} - {} wa_report.xlsx'.format(week_beginning,  # The format is: dd.mm - dd.mm
                                                   week_end)

    print('Saving report as {}...'.format(filename))
    wb.save(os.path.join(PATH_TO_REPORTS, filename))
