#! python3
#
# NAME         : change_date.py
#
# DESCRIPTION  : Changes the imported module name in work_attendance.py
#                according to the current date.
#
# AUTHOR       : Tim Kornev (Timmate profile on GitHub)
#
# CREATED DATE : 13th of July, 2016
#


import re
import datetime


FILENAME = 'work_attendance.py'

today_dt = datetime.datetime.now()
today_strf = today_dt.strftime('%d_%b_%Y')
new_module_name = 'wa_{}'.format(today_strf)
print()
print('TODAY: {}'.format(today_strf))

module_regex = re.compile(r'wa_\d{1,2}_[a-zA-Z]{3}_20\d\d')

print('OPEN: {}'.format(FILENAME))
read_file = open(FILENAME)
contents = read_file.read()
read_file.close()

# Find the module name and change it.
mo = module_regex.search(contents)  # 'mo' stands for Match Objects.
print('FOUND: {} \nCHANGE_TO: {}'.format(mo.group(0), new_module_name)) 
contents = module_regex.sub(new_module_name, contents)

write_file = open(FILENAME, 'w')
write_file.write(contents)
write_file.close()

print()
print('Done.')
print()
