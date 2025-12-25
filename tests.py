"""
Unit tests for the Attendance Processor
"""
import unittest
from datetime import time
from attendance_processor import AttendanceProcessor


class TestAttendanceProcessor(unittest.TestCase):
    """Test cases for AttendanceProcessor"""

    def setUp(self):
        """Set up test fixtures"""
        self.processor = AttendanceProcessor()
        self.processor.read_attendance_file('day1.txt', 'day1')
        self.processor.read_attendance_file('day2.txt', 'day2')

    # ==================== Time Conversion Tests ====================
    def test_time_to_minutes_morning(self):
        """Test conversion of morning time to minutes"""
        t = time(9, 30)
        expected = 9 * 60 + 30  # 570 minutes
        result = self.processor.time_to_minutes(t)
        self.assertEqual(result, expected)

    def test_time_to_minutes_midnight(self):
        """Test conversion of midnight to minutes"""
        t = time(0, 0)
        expected = 0
        result = self.processor.time_to_minutes(t)
        self.assertEqual(result, expected)

    def test_time_to_minutes_evening(self):
        """Test conversion of evening time to minutes"""
        t = time(20, 37)
        expected = 20 * 60 + 37  # 1237 minutes
        result = self.processor.time_to_minutes(t)
        self.assertEqual(result, expected)

    # ==================== Duration Calculation Tests ====================
    def test_calculate_duration_same_day(self):
        """Test duration calculation within same day"""
        entry = time(9, 0)
        exit_time = time(17, 0)
        expected = 8 * 60  # 480 minutes = 8 hours
        result = self.processor.calculate_duration(entry, exit_time)
        self.assertEqual(result, expected)

    def test_calculate_duration_cross_midnight(self):
        """Test duration calculation spanning midnight"""
        entry = time(20, 37)
        exit_time = time(0, 0)
        # (24*60 - 1237) + 0 = 1440 - 1237 = 203 minutes
        expected = 203
        result = self.processor.calculate_duration(entry, exit_time)
        self.assertEqual(result, expected)

    def test_calculate_duration_early_morning(self):
        """Test duration from evening to early morning"""
        entry = time(22, 0)
        exit_time = time(6, 0)
        # (24*60 - 1320) + 360 = 120 + 360 = 480 minutes
        expected = 480
        result = self.processor.calculate_duration(entry, exit_time)
        self.assertEqual(result, expected)

    # ==================== Break Deduction Tests ====================
    def test_deduct_breaks_within_shift(self):
        """Test break deduction for times within break periods"""
        entry = time(10, 30)
        exit_time = time(11, 15)
        breaks = [
            (time(10, 45), time(11, 0)),
            (time(15, 45), time(16, 0))
        ]
        # Break from 10:45 to 11:00 (15 mins) overlaps with 10:30-11:15
        result = self.processor.deduct_breaks(entry, exit_time, breaks)
        self.assertEqual(result, 15)

    def test_deduct_breaks_no_overlap(self):
        """Test break deduction when breaks don't overlap with shift"""
        entry = time(9, 0)
        exit_time = time(10, 0)
        breaks = [
            (time(10, 45), time(11, 0)),
            (time(15, 45), time(16, 0))
        ]
        result = self.processor.deduct_breaks(entry, exit_time, breaks)
        self.assertEqual(result, 0)

    def test_deduct_breaks_multiple_breaks(self):
        """Test break deduction with multiple breaks"""
        entry = time(9, 0)
        exit_time = time(20, 30)
        breaks = [
            (time(10, 45), time(11, 0)),  # 15 mins
            (time(15, 45), time(16, 0))   # 15 mins
        ]
        # Both breaks should be deducted
        result = self.processor.deduct_breaks(entry, exit_time, breaks)
        self.assertEqual(result, 30)  # 15 + 15

    # ==================== Grace Time Tests ====================
    def test_apply_grace_time_late_coming(self):
        """Test grace time for late coming (within 10 mins)"""
        entry = time(9, 5)  # 5 mins late
        exit_time = time(17, 0)
        shift_config = {'begin_time': time(9, 0), 'end_time': time(18, 0), 'breaks': []}
        rules = {'grace_late_coming': 10, 'grace_early_going': 10}

        result = self.processor.apply_grace_time(entry, exit_time, shift_config, rules)
        # Should adjust entry to 9:00, giving 8 hours exactly
        expected = 8 * 60
        self.assertEqual(result, expected)

    def test_apply_grace_time_no_grace_needed(self):
        """Test when grace time doesn't apply (on time)"""
        entry = time(9, 0)
        exit_time = time(17, 0)
        shift_config = {'begin_time': time(9, 0), 'end_time': time(18, 0), 'breaks': []}
        rules = {'grace_late_coming': 10, 'grace_early_going': 10}

        result = self.processor.apply_grace_time(entry, exit_time, shift_config, rules)
        expected = 8 * 60
        self.assertEqual(result, expected)

    # ==================== Cross-Midnight Shift Detection Tests ====================
    def test_cross_midnight_shift_detection(self):
        """Test detection of cross-midnight shift"""
        emp_code = 'EW00029'
        emp = self.processor.employees[emp_code]

        day1_adj, day2_adj, cross_info = self.processor.detect_cross_midnight_shift(
            emp['day1_punches'], emp['day2_punches']
        )

        # Should detect a cross-midnight shift
        self.assertIsNotNone(cross_info)
        self.assertTrue(cross_info.get('is_cross_midnight', False))
        self.assertEqual(cross_info['duration_mins'], 203)  # 20:37 to 00:00

    def test_cross_midnight_shift_hour_calculation(self):
        """Test calculation of cross-midnight shift hours"""
        entry = time(20, 37)
        exit_time = time(0, 0)

        result = self.processor.calculate_cross_midnight_hours(entry, exit_time)
        # 203 minutes - no breaks = 203 minutes = 3.38 hours
        expected = 203
        self.assertEqual(result, expected)

    # ==================== Employee Data Tests ====================
    def test_employee_ew00029_day1_hours(self):
        """Test that EW00029 Day 1 hours are calculated correctly"""
        emp = self.processor.employees['EW00029']

        # Adjust for cross-midnight
        day1_adj, day2_adj, cross_info = self.processor.detect_cross_midnight_shift(
            emp['day1_punches'], emp['day2_punches']
        )

        day1_hours = self.processor.calculate_working_hours(day1_adj)
        expected = int(10.57 * 60)  # 10.57 hours = 634 minutes

        # Allow small tolerance due to grace time and break calculations
        self.assertAlmostEqual(day1_hours, expected, delta=5)

    def test_employee_ew00029_day2_hours(self):
        """Test that EW00029 Day 2 hours are calculated correctly"""
        emp = self.processor.employees['EW00029']

        # Don't adjust Day 2 for cross-midnight (keep original punches)
        day2_hours = self.processor.calculate_working_hours(emp['day2_punches'])
        expected = int(20.65 * 60)  # 20.65 hours = 1239 minutes

        # Allow small tolerance due to grace time and break calculations
        self.assertAlmostEqual(day2_hours, expected, delta=5)

    def test_employee_ew00029_total_hours(self):
        """Test total hours for EW00029 across both days"""
        emp = self.processor.employees['EW00029']

        # Get adjusted punches
        day1_adj, day2_adj, cross_info = self.processor.detect_cross_midnight_shift(
            emp['day1_punches'], emp['day2_punches']
        )

        day1_hours = self.processor.calculate_working_hours(day1_adj)
        day2_hours = self.processor.calculate_working_hours(day2_adj)
        cross_hours = 203 if cross_info else 0  # 3.38 hours

        total = day1_hours + day2_hours + cross_hours
        expected = int(34.60 * 60)  # 34.60 hours = 2076 minutes

        self.assertAlmostEqual(total, expected, delta=5)

    # ==================== Excel Output Tests ====================
    def test_excel_file_generated(self):
        """Test that Excel file is generated"""
        import os

        output_file = 'test_output.xls'
        self.processor.generate_excel(output_file)

        # Check file exists
        self.assertTrue(os.path.exists(output_file))

        # Clean up
        if os.path.exists(output_file):
            os.remove(output_file)

    def test_all_employees_processed(self):
        """Test that all employees are present in the data"""
        # Should have at least 50 employees
        self.assertGreaterEqual(len(self.processor.employees), 50)

    def test_employee_has_required_fields(self):
        """Test that each employee has required data fields"""
        emp_code = 'EW00029'
        emp = self.processor.employees[emp_code]

        required_fields = ['name', 'company', 'department', 'day1_punches', 'day1_hours', 'day2_punches', 'day2_hours']
        for field in required_fields:
            self.assertIn(field, emp, f"Employee missing field: {field}")


class TestMinuteConversion(unittest.TestCase):
    """Test minute to decimal hour conversion"""

    def setUp(self):
        self.processor = AttendanceProcessor()

    def test_minutes_to_decimal_hours_exact(self):
        """Test conversion of exact hour values"""
        # 60 minutes = 1 hour
        result = self.processor.minutes_to_decimal_hours(60)
        self.assertEqual(result, 1.0)

    def test_minutes_to_decimal_hours_partial(self):
        """Test conversion of partial hours"""
        # 90 minutes = 1.5 hours
        result = self.processor.minutes_to_decimal_hours(90)
        self.assertEqual(result, 1.5)

    def test_minutes_to_decimal_hours_zero(self):
        """Test conversion of zero minutes"""
        result = self.processor.minutes_to_decimal_hours(0)
        self.assertEqual(result, 0.0)


if __name__ == '__main__':
    unittest.main(verbosity=2)
