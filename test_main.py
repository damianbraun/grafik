import unittest
from main import Shift
from datetime import date, datetime


class ShiftCase(unittest.TestCase):

    def test_datetime(self):
        shiftd = Shift('d', date(year=2017, month=3, day=4))
        self.assertEqual(shiftd.get_start_time(), datetime(year=2017, month=3,
                                                           day=4, hour=7))
        self.assertEqual(shiftd.get_end_time(), datetime(year=2017, month=3,
                                                         day=4, hour=19))

        shiftn = Shift('N', date(year=2020, month=1, day=24))
        self.assertEqual(shiftn.get_start_time(), datetime(year=2020, month=1,
                                                           day=24, hour=19))
        self.assertEqual(shiftn.get_end_time(), datetime(year=2020, month=1,
                                                         day=25, hour=7))

        shifto = Shift('ó', date(year=2013, month=12, day=6))
        self.assertEqual(shifto.get_start_time(), datetime(year=2013, month=12,
                                                           day=6, hour=7))
        self.assertEqual(shifto.get_end_time(), datetime(year=2013, month=12,
                                                         day=6, hour=15))


if __name__ == '__main__':
    unittest.main()
