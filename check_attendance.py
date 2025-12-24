import csv
from datetime import datetime, time
from collections import defaultdict
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ---------------- SHIFT CONFIG ---------------- #

SHIFTS = {
    'Shift': {
        'begin_time': time(9, 0),
        'end_time': time(20, 30),
        'breaks': [
            (time(10, 45), time(11, 0)),
            (time(13, 0), time(13, 30)),
            (time(15, 45), time(16, 0)),
            (time(20, 30), time(21, 0))
        ]
    }
}

CATEGORY_RULES = {
    'Cont Worked': {
        'deduct_break_hours': True
    }
}

# ---------------- PROCESSOR ---------------- #

class AttendanceProcessor:

    def __init__(self):
        self.employees = defaultdict(lambda: {
            'name': '',
            'company': '',
            'department': '',
            'day1_punches': [],
            'day2_punches': [],
            'day1_hours': 0,
            'day2_hours': 0,
            'total_hours': 0
        })

    # ---------- TIME HELPERS ---------- #

    def parse_time(self, s):
        try:
            return datetime.strptime(s.strip(), "%H:%M").time()
        except:
            return None

    def time_to_minutes(self, t):
        return t.hour * 60 + t.minute

    def calculate_duration(self, start, end):
        s = self.time_to_minutes(start)
        e = self.time_to_minutes(end)
        return e - s if e >= s else (1440 - s + e)

    def deduct_breaks(self, start, end, breaks):
        s = self.time_to_minutes(start)
        e = self.time_to_minutes(end)
        if e < s:
            e += 1440

        total = 0
        for b1, b2 in breaks:
            bs = self.time_to_minutes(b1)
            be = self.time_to_minutes(b2)
            if bs < e and be > s:
                total += max(0, min(be, e) - max(bs, s))
        return total

    def minutes_to_decimal(self, mins):
        return round(mins / 60, 2)

    # ---------- CORE CALC ---------- #

    def calculate_working_hours_total_span(self, punches, day):
        if len(punches) < 2:
            return 0

        punches = sorted(set(punches), key=self.time_to_minutes)

        if day == 'day1':
            punches = [p for p in punches if self.time_to_minutes(p) >= 540]

        if len(punches) < 2:
            return 0

        start, end = punches[0], punches[-1]
        duration = self.calculate_duration(start, end)

        duration -= self.deduct_breaks(start, end, SHIFTS['Shift']['breaks'])
        return max(0, duration)

    # ---------- MIDNIGHT MERGE ---------- #

    def merge_midnight_punches(self, day1, day2):
        boundary = self.time_to_minutes(time(6, 45))
        merged = list(day1)
        remaining = []

        for p in day2:
            if self.time_to_minutes(p) <= boundary:
                merged.append(p)
            else:
                remaining.append(p)

        return merged, remaining

    # ---------- FILE READ ---------- #

    def read_attendance_file(self, filepath, day_key):
        with open(filepath, 'r', encoding='utf-8') as f:
            reader = csv.reader(f)
            next(reader)

            for row in reader:
                if len(row) < 8:
                    continue

                emp = row[0].strip()
                if not emp or row[7] == 'Not Present':
                    continue

                self.employees[emp]['name'] = row[1]
                self.employees[emp]['company'] = row[2]
                self.employees[emp]['department'] = row[3]

                punches = []
                for p in row[6].split(','):
                    t = self.parse_time(p)
                    if t:
                        punches.append(t)

                self.employees[emp][f'{day_key}_punches'] = punches

    # ---------- MAIN PROCESS ---------- #

    def process(self, day1_file, day2_file, output):
        self.read_attendance_file(day1_file, 'day1')
        self.read_attendance_file(day2_file, 'day2')

        for emp in self.employees.values():
            merged, remaining = self.merge_midnight_punches(
                emp['day1_punches'], emp['day2_punches']
            )

            emp['day1_punches'] = merged
            emp['day2_punches'] = remaining

            emp['day1_hours'] = self.calculate_working_hours_total_span(merged, 'day1')
            emp['day2_hours'] = self.calculate_working_hours_total_span(remaining, 'day2')

            emp['total_hours'] = emp['day1_hours'] + emp['day2_hours']

        self.generate_excel(output)

    # ---------- EXCEL ---------- #

    def generate_excel(self, output):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Attendance"

        headers = [
            "Employee Code", "Name", "Company", "Department",
            "Day1 Punches", "Day1 Hours",
            "Day2 Punches", "Day2 Hours", "Total Hours"
        ]

        ws.append(headers)

        for emp_code, emp in self.employees.items():
            if emp['total_hours'] == 0:
                continue

            ws.append([
                emp_code,
                emp['name'],
                emp['company'],
                emp['department'],
                ", ".join(t.strftime("%H:%M") for t in emp['day1_punches']),
                self.minutes_to_decimal(emp['day1_hours']),
                ", ".join(t.strftime("%H:%M") for t in emp['day2_punches']),
                self.minutes_to_decimal(emp['day2_hours']),
                self.minutes_to_decimal(emp['total_hours'])
            ])

        wb.save(output)
        print("Excel generated:", output)

# ---------------- MAIN ---------------- #

if __name__ == "__main__":
    AttendanceProcessor().process(
        r"c:\contract\day1.txt",
        r"c:\contract\day2.txt",
        r"c:\contract\Desiredoutput_final_new.xlsx"
    )