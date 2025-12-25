from datetime import time
from attendance_processor import AttendanceProcessor, SHIFTS, CATEGORY_RULES

# Create processor instance
processor = AttendanceProcessor()

# Parse the punch times
punch_str = "00:01, 00:01, 08:57, 08:57, 10:45, 10:45, 11:00, 11:00, 12:57, 13:33, 13:33, 15:45, 15:45, 15:59, 15:59, 18:00, 18:00, 18:12, 18:12, 20:02, 20:36"
punches = processor.parse_punch_records(punch_str)

print("=" * 80)
print("PUNCH TIME LOG ANALYSIS")
print("=" * 80)
print(f"\nTotal punches: {len(punches)}")
print("\nPunch pairs (IN → OUT):")
print("-" * 80)

total_minutes = 0
shift_config = SHIFTS['Shift']
rules = CATEGORY_RULES['Cont Worked']

for i in range(0, len(punches) - 1, 2):
    entry = punches[i]
    exit_time = punches[i + 1]

    # Calculate raw duration
    duration = processor.calculate_duration(entry, exit_time)

    # Deduct breaks
    break_mins = processor.deduct_breaks(entry, exit_time, shift_config['breaks'])
    duration_after_break = duration - break_mins

    # Apply grace time
    grace_adjusted = processor.apply_grace_time(entry, exit_time, shift_config, rules)
    if grace_adjusted > 0:
        duration_final = grace_adjusted
        # Re-deduct breaks after grace
        break_mins = processor.deduct_breaks(entry, exit_time, shift_config['breaks'])
        duration_final -= break_mins
    else:
        duration_final = duration_after_break

    if duration_final > 0:
        total_minutes += duration_final
        print(f"Pair {i//2 + 1}: {entry} → {exit_time} = {duration} mins - {break_mins} mins breaks = {duration_final} mins ({duration_final/60:.2f} hrs)")
    else:
        print(f"Pair {i//2 + 1}: {entry} → {exit_time} = 0 mins (same punch)")

# Check if there's an unpaired punch at the end
if len(punches) % 2 == 1:
    print(f"\nUnpaired punch at end: {punches[-1]} (ignored)")

print("\n" + "=" * 80)
print(f"TOTAL WORKING TIME: {total_minutes} minutes = {total_minutes/60:.2f} hours")
print("=" * 80)

# Additional breakdown
print(f"\nBreakdown:")
print(f"  Raw minutes: {total_minutes}")
print(f"  Decimal hours: {processor.minutes_to_decimal_hours(total_minutes):.2f} hrs")
