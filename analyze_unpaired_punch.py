from datetime import time
from attendance_processor import AttendanceProcessor, SHIFTS, CATEGORY_RULES

# Create processor instance
processor = AttendanceProcessor()

# Parse the punch times
punch_str = "00:01, 00:01, 08:57, 08:57, 10:45, 10:45, 11:00, 11:00, 12:57, 13:33, 13:33, 15:45, 15:45, 15:59, 15:59, 18:00, 18:00, 18:12, 18:12, 20:02, 20:36"
punches = processor.parse_punch_records(punch_str)

print("=" * 80)
print("UNPAIRED PUNCH ANALYSIS")
print("=" * 80)

print(f"\nTotal punches: {len(punches)}")
print(f"Last punch: {punches[-1]}")
print(f"Second to last punch: {punches[-2]}")

print("\n" + "-" * 80)
print("SCENARIO 1: The 20:36 is a CLOCK IN (entry)")
print("-" * 80)
print(f"Employee clocked IN at 20:36 but never clocked OUT")
print(f"This time cannot be counted (incomplete pair)")
print(f"Unaccounted time: Unknown (depends on when they actually left)")

print("\n" + "-" * 80)
print("SCENARIO 2: The 20:36 is a CLOCK OUT (exit)")
print("-" * 80)
print(f"The pair should be: 20:02 → 20:36")
print(f"Duration: {punches[-1].hour * 60 + punches[-1].minute - (punches[-2].hour * 60 + punches[-2].minute)} minutes")

shift_config = SHIFTS['Shift']
rules = CATEGORY_RULES['Cont Worked']

# Calculate as if 20:02 → 20:36 is a valid pair
entry = punches[-2]  # 20:02
exit_time = punches[-1]  # 20:36

duration = processor.calculate_duration(entry, exit_time)
break_mins = processor.deduct_breaks(entry, exit_time, shift_config['breaks'])
final_duration = duration - break_mins

print(f"Raw duration: {duration} mins")
print(f"Breaks deducted: {break_mins} mins (breaks occur 10:45-11:00 and 15:45-16:00)")
print(f"Net duration: {final_duration} mins ({final_duration/60:.2f} hrs)")

print("\n" + "=" * 80)
print("COMPARISON:")
print("=" * 80)
print(f"Current calculation (20:36 ignored): 410 minutes = 6.83 hours")
print(f"If 20:36 is exit time: {410 + final_duration} minutes = {(410 + final_duration)/60:.2f} hours")
print(f"Difference: {final_duration} minutes ({final_duration/60:.2f} hours)")

print("\n" + "=" * 80)
print("RECOMMENDATION:")
print("=" * 80)
print(f"The 20:36 punch is likely a CLOCK IN (entry) without a corresponding exit.")
print(f"This could indicate:")
print(f"  • Employee continued working after 20:02 but forgot to clock out")
print(f"  • System error or missing punch record")
print(f"  • Need to manually verify with employee or manager")
