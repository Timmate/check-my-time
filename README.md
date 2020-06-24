## check-my-time

A simple time and attendance system.

It is a simplified version of [easyclocking][1] service and some of [these][2]
time and attendance systems. You can read more on this on [Wiki page][3].

Python 3.x only. (for Python 2.x backward compatibility see [this][4])


### `check_my_time.py` usage

* Simply run this script wherever to start working. It will automatically create
  a tree of directories where all text files with data will be stored.

* In the app:
  * Enter a person's name and optional time argument for the first
    time a day to record time he/she clocked in and calculate time he/she was
    late for / early that day.

  * Enter the same name and optional time argument for the second time a day to record time he/she clocked out
    and calculate time he/she has worked for.

* Of course, you are able to close the app and load data already written for the
  current day later. Also, you can change start time of a working day in case the
  script is run for the first time a day.


### `report_creator.py` usage

(Please note that before using this script you need to install `openpyxl` module
that handles work with Excel spreadsheets.
One way to do that is to run `pip install -r requirements.txt`
in command line — this will prevent errors that may be caused by version
incompatibility.)

* Simply enter any date that belongs to a particular week or a month to create
  a report for that period of time.
  E.g. entering `M` (Month Report) and `16/06/2016` (it is Thursday) creates a month
  report for the whole June of 2016 (from 01/06/2016 to 30/06/2016).
  The same thing with `W` (Week Report), except that a report for one working
  week will be created (from 13/06/2016 to 19/06/2016).

* Also, you are able to choose kind of report: Simple or Complex. Simple Reports
  contain only overall time values, while Complex Reports also include average
  time values.

* You can test the script by running it on test data in *Test Data* directory or
  create your own text files with data with `check_my_time.py`.
  Note that *Report* and *Work Attendance Files* directories must always be in
  the same directory as `report_creator.py`.


[1]: http://easyclocking.com/
[2]: http://www.businessnewsdaily.com/6730-best-time-and-attendance-systems.html
[3]: https://en.wikipedia.org/wiki/Time_and_attendance
[4]: http://python-future.org/quickstart.html
