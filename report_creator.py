#! python3
#
# NAME          : report_creator.py
#
# DESCRIPTION   : Creates a work attendance report for a month or a week.
#
# AUTHOR        : Tim Kornev (@Timmate on GitHub)
#
# CREATED DATE  : 19th of July, 2016
#
# LAST MODIFIED : 29th of October, 2016
#


import os
import ast    # ast is used for parsing data from text files
import datetime

import openpyxl


# Constants.
# ==========

WORKING_DIR = 'Work Attendance Files'    # dir where text files are stored
REPORTS_DIR = 'Reports'    # dir where reports are stored
TEMPLATES_DIR = 'Templates'    # dir where spreadsheet templates are stored

DATE_FORMAT = 'dd/mm/yyyy'    # can be changed to American date format
DATE_FORMAT_STRPTIME = '%d/%m/%Y'    # if `DATE_FORMAT` was changed to American date format,
                                     # then this needs to be changed, too.
WORKDAYS_PER_WEEK = 5     # is used in average time values calculations
ZERO_PADDED_FILENAMES = False    # is related to text files created by `check_my_time.py`
SPREADSHEET_SAVE_EXTENSTION = 'xlsx'


# Functions.
# ==========


def main():
    """The core function."""

    # Ask the user whether to create a report or not.
    while True:
        print()

        choice = None
        while not choice in ('y', 'n'):
            choice = input("May I create a report for you? [y/n]: ")
            choice = choice.lower()

        if choice == 'y':
            # Create a report.
            get_date_input()
            parse_date_input()
            gather_data()
            calculate_time()
            write_to_spreadsheet()

        elif choice == 'n':
            # Exit.
            print()
            print('Goodbye.')
            print()

            break




def get_date_input():
    """Asks the user for date input."""

    # These globals are used by these functions: `parse_date_input()`, `write_to_spreadsheet()`,
    global month_or_week, report_complexity, date_d

    print()

    # Ask the user what time period the report should be created for.
    choice = None

    while not choice in ('M', 'W'):
        choice = input("Enter 'M' for Month or 'W' for Week to create a report: ")
        choice = choice.upper()

    # Save the user's choice.
    month_or_week = choice

    # Ask the user what kind of report should be created.
    choice = None

    while not choice in ('S', 'C'):
        choice = input("Enter 'S' for Simple or 'C' for complex report: ")
        choice = choice.upper()

    # Save the user's choice.
    report_complexity = choice

    print()

    # Ask for date input.
    while True:
        date_strf = input('Enter date ({}): '.format(DATE_FORMAT))

        try:
            # Parse the date input.
            date_dt = datetime.datetime.strptime(date_strf, DATE_FORMAT_STRPTIME)
            # We do not really need datetime objects, so convert it to a date object.
            date_d = datetime.date(date_dt.year, date_dt.month, date_dt.day)
        except ValueError as err:     # if the date input was incorrect
            print('Error: ' + str(err))
            continue
        else:
            break




def parse_date_input():
    """
    Parses the date input. Here, parsing means simply finding start and end dates
    of a period of time the date input belongs to.
    """

    # These globals are used by these functions: `gather_data()`, `calculate_time()`,
    # `write_to_spreadsheet()`
    global start, end

    if month_or_week == 'M':
        # Find first and last days of a month.
        year, month = date_d.year, date_d.month
        start = datetime.datetime(year, month, 1)   # get the first day

        try:
            # Get the last day of the month by subtracting one day from the first day of the next month.
            end = datetime.datetime(year, month+1, 1) - datetime.timedelta(days=1)
        except ValueError:
            # If we are dealing with December, ValueError is raised, saying that 'month must be in 1..12'.
            # Get the last day of the month by subtracting one day from the first day of January (of course,
            # it will be the 31st of December, but use the 'classic' method)
            end = datetime.datetime(year+1, 1, 1) - datetime.timedelta(days=1)

    elif month_or_week == 'W':
        # Find start and end of the week.
        weekday = date_d.weekday()
        start = date_d - datetime.timedelta(days=weekday)
        end = start + datetime.timedelta(days=6)




def gather_data():
    """
    Gathers data from text files and calculates sum of work, late, and early
    time values.
    """

    # These globals are used by these functions: `calculate_time()`, `write_to_spreadsheet()`
    global data_sum, days_counter, rel_path_to_month_dir, month_name

    # NOTE: in case all dirs are missing or exist but do not contain any files,
    # `write_to_spreadsheet()` will create an empty spreadsheet anyway.

    data_sum = {}    # stores data sum
    days_counter = 0    # this is important to calculating average time values
                        # in calculate_time() function.
    printed_dirs = []    # store dirs that are printed as missing or existing ones

    date = start    # begin from the first day

    while date <= end:
        # Find dirs' names according to date value.
        year_dir = date.strftime('%Y')
        month_dir = date.strftime('%-m — %B')
        month_name = date.strftime('%B')

        # Find filename according to date value.
        if ZERO_PADDED_FILENAMES:
            filename = date.strftime('%d') + '.txt'
        else:
            filename = date.strftime('%-d') + '.txt'

        # Find paths to dirs and filename.
        full_path_to_year_dir = os.path.join(WORKING_DIR, year_dir)
        full_path_to_month_dir = os.path.join(full_path_to_year_dir, month_dir)
        # 'rel' stands for 'relative'
        rel_path_to_month_dir = os.path.join(year_dir, month_dir)
        full_path_to_filename = os.path.join(full_path_to_month_dir, filename)

        if not year_dir in os.listdir(WORKING_DIR):    # if no year dir found
            if not year_dir in printed_dirs:    # if missing dir was not displayed
                print()
                print('! Missing directory: {}'.format(year_dir))
                # Add it to printed_dir so that it will not be displayed the next time.
                # This means that dirs are printed only once.
                # NOTE: the same principle applies to all printed dirs.
                printed_dirs.append(year_dir)

        else:     # if year dir found
            if not month_dir in os.listdir(full_path_to_year_dir):     # if no month dir in year dir found
                if not rel_path_to_month_dir in printed_dirs:
                    # Display missing dir.
                    print()
                    print('! Missing directory: {}'.format(rel_path_to_month_dir))
                    printed_dirs.append(rel_path_to_month_dir)

            else:    # if month dir in year dir found
                if not rel_path_to_month_dir in printed_dirs:
                    # Display existing dir.
                    print()
                    print('Looking into directory {} ...'.format(rel_path_to_month_dir))
                    printed_dirs.append(rel_path_to_month_dir)

                if not filename in os.listdir(full_path_to_month_dir):    # if no filename in month dir found
                    # Display missing filename.
                    print('\t' + '! Missing file {}'.format(filename))

                else:   # if filename in month dir found
                    # Gather data from the file.

                    # Display existing filename.
                    print('\t' + 'Reading file {} ...'.format(filename))

                    days_counter += 1    # increments each time a file is read

                    # Read file.
                    with open(full_path_to_filename) as f:
                        data = ast.literal_eval(f.read())

                    # Gather data for each person in the file.
                    for name in data:
                        if name == 'day_start':
                            # We do not need this value.
                            continue
                        else:
                            data_sum.setdefault(name, {'early_time_hour': 0,
                                                       'early_time_minute': 0,
                                                       'late_time_hour': 0,
                                                       'late_time_minute': 0,
                                                       'work_time_hour': 0,
                                                       'work_time_minute': 0
                                                       })


                            for key in data_sum[name]:
                                if key in data[name]:
                                    # Add new value to the existing one.
                                    value = eval(str(data[name][key]))
                                    data_sum[name][key] += value

        # Go to the next day.
        date += datetime.timedelta(days=1)




def calculate_time():
    """Calculates work, late, and early overall and average time values."""

    # This global is used by `write_to_spreadsheet()` function.
    global data

    data = {}    # stores calculated time values

    # Gather data from `data_sum` dict, make calculations, and write results to
    # `data` dict.
    for name in data_sum:
        data[name] = {}

        for category in ['early', 'late', 'work']:    # for each of categories
            # Calculate overall time value.
            hour = data_sum[name]['{}_time_hour'.format(category)]
            minute = data_sum[name]['{}_time_minute'.format(category)]
            hour, minute = clean_time(hour, minute)    # see `clean_time()` docstring for details
            # Write results to `data` dict.
            data[name]['{}_time_hour_overall'.format(category)] = hour
            data[name]['{}_time_minute_overall'.format(category)] = minute

            # Save overall values in order to use them in average time calculations.
            hour_overall, minute_overall = hour, minute

            if report_complexity == 'S':
                # Simple reports contain only overall time values, so continue.
                continue
            else:
                # Calculate average time values.

                # Calculate average time per day by dividing overall time values
                # by number of workdays. `days_counter` contains number of files
                # data was gathered from, so it could be considered the number
                # of workdays.
                hour =  hour_overall / days_counter
                minute = minute_overall / days_counter
                hour, minute = clean_time_2(hour, minute)
                data[name]['{}_time_hour_average_per_day'.format(category)] = hour
                data[name]['{}_time_minute_average_per_day'.format(category)] = minute

                if month_or_week == 'M':
                    # Also, calculate average time per week for month reports by
                    # dividing the number of workdays by a number of working
                    # weeks in the month.

                    # Calculate the number of working weeks in the month.
                    working_weeks = days_counter / WORKDAYS_PER_WEEK

                    hour = hour_overall / working_weeks
                    minute = minute_overall / working_weeks
                    hour, minute = clean_time_2(hour, minute)    # see `clean_time_2()` docstring for details
                    data[name]['{}_time_hour_average_per_week'.format(category)] = hour
                    data[name]['{}_time_minute_average_per_week'.format(category)] = minute




def clean_time(hour, minute):
    """
    Cleans time format. Since `gather_data()` finds separate sums of hours and
    minutes, 'extra' minutes should be converted into hours.
    (e.g. 10 hours, 180 minutes ---> 13 hours, 0 minutes)
    """

    if minute >= 60:
        hour += minute // 60
        minute %= 60

    return hour, minute




def clean_time_2(hour, minute):
    """
    Cleans time format 2. Since hours and minutes become float values after
    true division operations in `calculate_time()`, they should be formatted
    as follows:
    – hours value's decimal part shoud be converted into minutes;
    – 'extra' minutes should be converted into hours.
    (e.g. 10.5 hours, 198.12 minutes ---> 13 hours, 48 minutes)
    """

    decimal_minute = (hour - int(hour)) * 10
    minute += decimal_minute * 60 / 10
    hour = int(hour)
    minute = int(minute)

    if minute >= 60:
        hour += minute // 60
        minute %= 60

    return hour, minute




def write_to_spreadsheet():
    """Writes data to an Excel spreadsheet using a template."""

    # This global is needed to redefine `rel_path_to_month_dir` later.
    # Without this statement, redefinition will raise an error.
    global rel_path_to_month_dir

    if report_complexity == 'S':    # for simple report
        # Open a template.
        path_to_template = os.path.join(TEMPLATES_DIR, 'Simple.xlsx')
        wb = openpyxl.load_workbook(path_to_template)
        sheet = wb.active
        row = 8    # start from 8th row

        # Write kind of report with the number of workdays.
        if month_or_week == 'W':
            sheet['A1'] = 'Week Report ({} day(s))'.format(days_counter)

        elif month_or_week == 'M':
            sheet['A1'] = 'Month Report ({} day(s))'.format(days_counter)

        # Write data to cells according to template's structure.
        for name in sorted(data):    # note that names are written in alphabetic order

            sheet['A' + str(row)] = name
            sheet['B' + str(row)] = data[name]['work_time_hour_overall']
            sheet['C' + str(row)] = data[name]['work_time_minute_overall']
            sheet['D' + str(row)] = data[name]['late_time_hour_overall']
            sheet['E' + str(row)] = data[name]['late_time_minute_overall']
            sheet['F' + str(row)] = data[name]['early_time_hour_overall']
            sheet['G' + str(row)] = data[name]['early_time_minute_overall']

            row += 1


    elif report_complexity == 'C':    # for complex report
        # Open a template.
        path_to_template = os.path.join(TEMPLATES_DIR, 'Complex.xlsx')
        wb = openpyxl.load_workbook(path_to_template)

        if month_or_week == 'W':
            # Delete `Month Report` spreadsheet.
            wb.remove(wb.get_sheet_by_name('Month Report'))
            wb.active = 0    # this helps to avoid raising an error
            sheet = wb.get_sheet_by_name('Week Report')
            row = 9    # start from 9th row

            # Write kind of report with the number of workdays.
            sheet['A1'] = 'Week Report ({} day(s))'.format(days_counter)

            # Write data to cells according to template's structure.
            for name in sorted(data):

                sheet['A' + str(row)] = name
                sheet['B' + str(row)] = data[name]['work_time_hour_overall']
                sheet['C' + str(row)] = data[name]['work_time_minute_overall']
                sheet['D' + str(row)] = data[name]['work_time_hour_average_per_day']
                sheet['E' + str(row)] = data[name]['work_time_minute_average_per_day']
                sheet['F' + str(row)] = data[name]['late_time_hour_overall']
                sheet['G' + str(row)] = data[name]['late_time_minute_overall']
                sheet['H' + str(row)] = data[name]['late_time_hour_average_per_day']
                sheet['I' + str(row)] = data[name]['late_time_minute_average_per_day']
                sheet['J' + str(row)] = data[name]['early_time_hour_overall']
                sheet['K' + str(row)] = data[name]['early_time_minute_overall']
                sheet['L' + str(row)] = data[name]['early_time_hour_average_per_day']
                sheet['M' + str(row)] = data[name]['early_time_minute_average_per_day']

                row += 1

        elif month_or_week == 'M':
            # Delete `Week Report` spreadsheet.
            wb.remove(wb.get_sheet_by_name('Week Report'))
            wb.active = 0    # this helps to avoid raising an error
            sheet = wb.get_sheet_by_name('Month Report')
            row = 10    # start from 10th row

            # Write kind of report with the number of workdays.
            sheet['A1'] = 'Month Report ({} day(s), {} workday(s) per week)'.format(days_counter, WORKDAYS_PER_WEEK)

            # Write data to cells according to template's structure.
            for name in sorted(data):

                sheet['A' + str(row)] = name
                sheet['B' + str(row)] = data[name]['work_time_hour_overall']
                sheet['C' + str(row)] = data[name]['work_time_minute_overall']
                sheet['D' + str(row)] = data[name]['work_time_hour_average_per_day']
                sheet['E' + str(row)] = data[name]['work_time_minute_average_per_day']
                sheet['F' + str(row)] = data[name]['work_time_hour_average_per_week']
                sheet['G' + str(row)] = data[name]['work_time_minute_average_per_week']
                sheet['H' + str(row)] = data[name]['late_time_hour_overall']
                sheet['I' + str(row)] = data[name]['late_time_minute_overall']
                sheet['J' + str(row)] = data[name]['late_time_hour_average_per_day']
                sheet['K' + str(row)] = data[name]['late_time_minute_average_per_day']
                sheet['L' + str(row)] = data[name]['late_time_hour_average_per_week']
                sheet['M' + str(row)] = data[name]['late_time_minute_average_per_week']
                sheet['N' + str(row)] = data[name]['early_time_hour_overall']
                sheet['O' + str(row)] = data[name]['early_time_minute_overall']
                sheet['P' + str(row)] = data[name]['early_time_hour_average_per_day']
                sheet['Q' + str(row)] = data[name]['early_time_minute_average_per_day']
                sheet['R' + str(row)] = data[name]['early_time_hour_average_per_week']
                sheet['S' + str(row)] = data[name]['early_time_minute_average_per_week']

                row += 1

    # Save the spreadsheet.
    rel_path_to_month_dir = os.path.join(REPORTS_DIR, rel_path_to_month_dir)
    os.makedirs(rel_path_to_month_dir, exist_ok=True)

    if month_or_week == 'W':

        # Use separate dirs for simple and complex reports to avoid name conflicts.
        if report_complexity == 'S':
            week_dir = 'Week Reports (Simple)'

        elif report_complexity == 'C':
            week_dir = 'Week Reports (Complex)'

        path_to_week_dir = os.path.join(rel_path_to_month_dir, week_dir)
        os.makedirs(path_to_week_dir, exist_ok=True)

        # Use start and end days in week report's name.

        # Apply filenames format to spreadsheet's name. Assume that it is the
        # format the user wants to use to name hir/her spreadsheets.
        if ZERO_PADDED_FILENAMES:
            start_strf = start.strftime('%d')
            end_strf = end.strftime('%d')
        else:
            start_strf = start.strftime('%-d')
            end_strf = end.strftime('%-d')

        # Spreadsheet's save extension is defined on the top of the script.
        # Note that if week starts in one month and ends in another one, it will
        # be saved to `end`s month dir.
        spreadsheet_name = '{}—{}.{}'.format(start_strf, end_strf, SPREADSHEET_SAVE_EXTENSTION)
        path_to_spreadsheet = os.path.join(path_to_week_dir, spreadsheet_name)
        wb.save(path_to_spreadsheet)

    elif month_or_week == 'M':

        if report_complexity == 'S':
            spreadsheet_name = 'Month Report for {} (Simple).{}'.format(month_name, SPREADSHEET_SAVE_EXTENSTION)
        elif report_complexity == 'C':
            spreadsheet_name = 'Month Report for {} (Complex).{}'.format(month_name, SPREADSHEET_SAVE_EXTENSTION)

        path_to_spreadsheet = os.path.join(rel_path_to_month_dir, spreadsheet_name)
        wb.save(path_to_spreadsheet)

    # Display name of the saved file and path to it.
    head, tail = os.path.split(path_to_spreadsheet)
    print()
    print('Saved as {} to {}.'.format(tail, head))


if __name__ == '__main__':
    main()
