from datetime import time
from attendance_processor import AttendanceProcessor

processor = AttendanceProcessor()

# Duration from 8:50 AM to 00:01 (12:01 AM)
entry = time(8, 50)
exit_time = time(0, 1)

duration_mins = processor.calculate_duration(entry, exit_time)
duration_hrs = duration_mins / 60

print("=" * 60)
print("DURATION CALCULATION")
print("=" * 60)
print(f"\nFrom: {entry} (8:50 AM)")
print(f"To:   {exit_time} (12:01 AM - next day)")
print(f"\nDuration: {duration_mins} minutes = {duration_hrs:.2f} hours")
print(f"          {int(duration_hrs)} hours {duration_mins % 60} minutes")
print("=" * 60)
