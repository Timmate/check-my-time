## work-attendance
A couple of scripts that helps you become a King (or Queen) of Pedantry.

Actually, it was a small and neat homework project for [Al Sweigart's](https://inventwithpython.com/) amazing [Automate the Boring Stuff with Python](https://automatetheboringstuff.com/chapter15/) book
that somehow grew into this. I was so encouraged to experiment with the `datetime` module
and try to create something helpful and useful that I eventually came up with this
simple and free app that is similar to [easyclocking](http://easyclocking.com/) service or
some of [these](http://www.businessnewsdaily.com/6730-best-time-and-attendance-systems.html)
time and attendance systems.

## `work_attendance.py` usage
Firstly, run `change_date.py` to change the imported module name in the code of this script.
You should do it only once a day, in the morning, when you start using this app. Otherwise,
the script will not work correctly.
When you are ready, run `work_attendance.py`. Simply enter the person's name and optional time argument for the first time a day to record the time he/she clocked in and how many minutes he/she was late today.
Enter the same name for the second time a day to record the time he/she clocked out and compute their working time.
Note that you are able to close the app and load data already written for the current day
later, when you decide to continue working. There is also an opportunity to change start time of the working day if the script is run for the first time a day.

## `change_date.py` usage
Just run it in the Python Shell or IDLE's Python interpreter to change the module name
that is used by `work_attendance.py`.

## `report.py` usage
(Please note that before using this script you need to install `openpyxl` module
that handles the work with Excel spreadsheets. Use the `pip` utility for that.)
Copy some files from *test_data* to *wa_files* or create your own files with data
there to test this script.
Entering `MONTH` and `16/06/2016` (it is Thursday) creates a month report for
the whole June of 2016 (from the 01/06/2016 to 30/06/2016). The same thing
with `WEEK`, except that a report for one working week will be created (from
13/06/2016 to 19/06/2016). You could enter any date that belongs to the
working month or the week to create a report for that period of time.
Note that days values in `'wa_...'` files must not be a zero-padded
decimal number.

## `functions.py` usage
You do not need to use this script as it just contains the functions that are used by `report.py`.

## License
Copyright © 2016 Tim Kornev (@Timmate).

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

http://www.apache.org/licenses/LICENSE-2.0

See the License file included in this repository for further details.
