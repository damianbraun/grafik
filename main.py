# coding=utf-8
import xlrd
import datetime
import calendar
from icalendar import Calendar, Event, vDatetime
import sys
import re
import logging
from difflib import SequenceMatcher

logging.basicConfig(filename='main.log', level=logging.DEBUG)
EMPLOYEE_NUM_COL = 0
EMPLOYEE_NAME_COL = 1
EMPLOYEE_EXCLUDE = ['nominał', 'L.p.']
POLISH_MONTH_LIST = ['Styczeń', 'Luty', 'Marzec', 'Kwiecień', 'Maj',
                     'Czerwiec', 'Lipiec', 'Sierpień', 'Wrzesień',
                     'Październik', 'Listopad', 'Grudzień']
DATE_COL_ROW = (11, 13)
FIRST_SHIFT_COL = 3
REG_SHIFT_DURATION = 12
SHORT_SHIFT_DURATION = 8
WORK_LOCATION = ''


class ShiftTimeException(Exception):
    pass


class Schedule():
    def __init__(self, filepath):
        self.filepath = filepath
        self.employees = []
        self.load_file()

    def load_file(self):
        '''Load spreadsheet file.'''
        book = xlrd.open_workbook(self.filepath)
        self.sh = book.sheet_by_index(0)

    def parse(self):
        '''Parse date and employee data from spreadsheet.'''
        # Extracting month an year form spreadsheet
        text = self.sh.cell(*DATE_COL_ROW).value  # Extracting from spreadsheet
        matchObj = re.match(r'(\w+).*(\d{4})', text)
        monthname = matchObj.group(1)
        year = int(matchObj.group(2))

        # Matching month name to month number
        lista = []
        month_num = 1
        for month in POLISH_MONTH_LIST:
            ratio = SequenceMatcher(None, monthname, month).ratio()
            lista.append((ratio, month, month_num))
            month_num += 1
        lista.sort(reverse=True)  # Sorting by similarity
        self.date = datetime.date(year=year, month=lista[0][2], day=1)

        # Calculating days count for two months
        self.days_count = calendar.monthrange(self.date.year,
                                              self.date.month)[1]
        + calendar.monthrange(self.date.year, self.date.month + 1)[1]

        # Add rows with data to list (exclude header)
        employees = []
        for i in range(self.sh.nrows):
            name = self.sh.col_values(1)[i]
            if name in EMPLOYEE_EXCLUDE:
                continue
            num = self.sh.col_values(0)[i]
            if num in EMPLOYEE_EXCLUDE:
                continue
            if name or num:
                employees.append([i, name, num])

        # Workaround for merged cells
        for index, employee in enumerate(employees):
            if employee[0] % 2 == 1:
                employee[2] = employees[index - 1][2]
                employee[0] = employee[0] - 1
                employees.pop(index - 1)

        # Conversion to int
        for employee in employees:
            try:
                employee[2] = int(employee[2])
            except ValueError:
                logging.info('Not a number in L.p. column')
                pass

        # Creating Employee objects for every employee
        for e in employees:
            self.add_employee(Employee(name=e[1], number=e[2], row=e[0],
                                       sh=self.sh, date=self.date,
                                       days=self.days_count))
        self.print_employees()

    def add_employee(self, employee):
        '''Add Employee, checking for validity.'''
        if isinstance(employee, Employee):
            self.employees.append(employee)
        else:
            raise TypeError('Not a Employee instance passed.')

    def export(self):
        '''Exports employees schedules to files.'''
        for employee in self.employees:
            employee.to_calendar()

    def print_employees(self):
        '''Prints out basic employees info.'''
        for employee in self.employees:
            print(employee)


class Employee():
    day = datetime.timedelta(days=1)

    def __init__(self, name, number, row, sh, date, days):
        self.name = name
        self.number = number
        self.row = row
        self.sh = sh
        self.date = date
        self.days = days
        self.shifts = []
        self.parse()

    def __str__(self):
        return '{}. {} row {}, {} shifts'.format(self.number, self.name,
                                                 self.row, len(self.shifts))

    def add_shift(self, shift):
        '''Adds shift, checking for validity'''
        if isinstance(shift, Shift):
            self.shifts.append(shift)
        else:
            raise TypeError('Not a Shift instance passed.')

    def parse(self):
        '''Parse row with shifts for one Employee.'''
        cells = self.sh.row_values(self.row)[FIRST_SHIFT_COL:]
        i = 0
        while i < self.days:
            if cells[i]:
                self.add_shift(Shift(time=cells[i], date=self.date))
            self.date = self.date + self.day
            i += 1

    def to_calendar(self):
        '''Exports shifts to iCalendar file.'''
        if self.shifts:
            self.calendar = Calendar()
            for shift in self.shifts:
                event = Event()
                event['summary'] = 'Zmiana {}'.format(shift.time)
                event['dtstart'] = vDatetime(shift.get_start_time()).to_ical()
                event['dtend'] = vDatetime(shift.get_end_time()).to_ical()
                event['location'] = WORK_LOCATION
                self.calendar.add_component(event)
            print(self.calendar.to_ical())
            file = open('{}.ics'.format(self.number), 'wb')
            file.write(self.calendar.to_ical())
            file.close()
        else:
            logging.info('No shifts for Employee({})'.format(self))


class Shift():
    shift_time_d = datetime.time(hour=7)
    shift_time_n = datetime.time(hour=19)

    def __init__(self, time, date):
        # Check type of shift
        if time in ('D', 'N', 'Ó', 'd', 'n', 'ó'):
            self.time = time.upper()
        elif time == '':
            return
        else:
            logging.info('Invalid time format, only \'D\', \'N\' and \'Ó\' is accepted.')
        # Date instance check
        if isinstance(date, datetime.date):
            self.date = date
        else:
            raise TypeError('Not a date instance passed.')

    def __str__(self):
        return '{} shift, from {} to {}'.format(self.time,
                                                self.get_start_time(),
                                                self.get_end_time())

    def get_start_time(self):
        '''Returns datetime object of shift start depending on Shift.time'''
        if self.time == 'D':
            start_dt = datetime.datetime.combine(self.date, self.shift_time_d)
        elif self.time == 'N':
            start_dt = datetime.datetime.combine(self.date, self.shift_time_n)
        elif self.time == 'Ó':
            start_dt = datetime.datetime.combine(self.date, self.shift_time_d)
        return start_dt

    def get_end_time(self):
        '''Returns datetime object of shift end depending on Shift.time'''
        if self.time == 'Ó':
            return self.get_start_time() + datetime.timedelta(hours=SHORT_SHIFT_DURATION)
        else:
            return self.get_start_time() + datetime.timedelta(hours=REG_SHIFT_DURATION)


def main():
    if len(sys.argv) > 1:
        logging.info('Program started.')
        schedule = Schedule(sys.argv[1])
        schedule.parse()
        schedule.export()
        logging.info('Program ended.')
    else:
        logging.info('No filepath specified.')
        print('Specify filepath.')


if __name__ == '__main__':
    main()
