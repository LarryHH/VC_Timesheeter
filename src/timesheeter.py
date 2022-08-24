from os import mkdir
from datetime import datetime, timedelta

MONTHS = [
    '01 JAN',
    '02 FEB',
    '03 MAR',
    '04 APR',
    '05 MAY',
    '06 JUN',
    '07 JUL',
    '08 AUG',
    '09 SEP',
    '10 OCT',
    '11 NOV',
    '12 DEC'
]


class Timesheeter():
    """ timesheeter class """

    def __init__(self, args) -> None:

        self.ERRORS = []

        self.wb = args[0]
        self.fp_template = args[1]
        self.year = int(args[2])
        self.first_monday = datetime.strptime(args[3], "%Y-%M-%d").date()
        self.date_sheet = args[4]
        self.date_cell = args[5]
        self.fp_output = args[6]

        self.done = False

    def construct_directories(self):
        """make folders for all months in year"""
        try:
            mkdir(self.fp_output)
        except FileExistsError as e:
            self.ERRORS.append('FileExistsError')
        except FileNotFoundError as e:
            self.ERRORS.append('FileNotFoundError')
        for month in MONTHS:
            try:
                mkdir(f"{self.fp_output}/{month}")
            except OSError as e:
                self.ERRORS.append('OSError')

    def excel_name(self, date, index):
        """format excel filename based on timesheet range"""
        month_now = str(date.strftime("%B"))[:3].upper()
        day_now = date.day
        date_fortnight = date + timedelta(days=13)
        month_fortnight = str(date_fortnight.strftime("%B"))[:3].upper()
        day_fortnight = date_fortnight.day
        index = f"0{index}"
        return f"{index} {day_now} {month_now} - {day_fortnight} {month_fortnight}.xlsx"

    def excel_write(self, date, name):
        """write excel workbooks for each timesheet period"""
        for month in MONTHS:
            if str(date.strftime("%B"))[:3].upper() in month:
                self.wb.save(filename=f"{self.fp_output}/{month}/{name}")

    def date_year_calc(self, offset):
        """calculate valid date ranges for given year"""
        monday = 0
        months = range(1, 13)
        days = range(1, 32)

        sheet_date = self.wb[self.date_sheet]

        fortnight = 1 if offset else 0
        for month in months:
            index = 1
            for day in days:
                try:
                    if datetime(self.year, month, day).weekday() == monday:
                        fortnight += 1
                        if fortnight % 2 == 0:
                            date = datetime(self.year, month, day)
                            sheet_date[self.date_cell] = date
                            name = self.excel_name(date, index)
                            print(name)
                            self.excel_write(date, name)
                            index += 1
                except Exception:  # day out of range ( < 31)
                    continue
        return None

    def is_good_to_go(self):
        print('attempting to execute...')
        self.construct_directories()
        return list(set(self.ERRORS))

    def exec(self):
        print('good to go!')
        offset = True
        if self.first_monday.day > 7:
            offset = False
        self.date_year_calc(offset)
        print('All done!')
        self.done = True
