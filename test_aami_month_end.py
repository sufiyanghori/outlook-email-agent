import datetime as dt
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock
from zoneinfo import ZoneInfo

import aami_month_end as job

TZ = ZoneInfo('Australia/Melbourne')


class MonthEndTests(unittest.TestCase):
    def test_not_due_before_first_month_end(self):
        self.assertEqual(job.due_months(dt.datetime(2026, 10, 31, 17, tzinfo=TZ)), [])

    def test_due_on_last_day_at_six(self):
        self.assertEqual(job.due_months(dt.datetime(2026, 10, 31, 18, tzinfo=TZ)), ['2026-10'])

    def test_catch_up_after_sleep(self):
        self.assertEqual(job.due_months(dt.datetime(2026, 11, 1, 9, tzinfo=TZ)), ['2026-10'])

    def test_leap_year_and_short_months(self):
        self.assertIn('2028-02', job.due_months(dt.datetime(2028, 2, 29, 18, tzinfo=TZ)))
        self.assertNotIn('2027-02', job.due_months(dt.datetime(2027, 2, 28, 17, tzinfo=TZ)))

    def test_timezone_conversion(self):
        now = dt.datetime(2026, 10, 31, 7, tzinfo=dt.timezone.utc)
        self.assertEqual(job.due_months(now), ['2026-10'])

    def test_message_requests_individual_receipts(self):
        subject, body = job.message('2026-10')
        self.assertIn('October 2026', subject)
        self.assertIn('62302014', body)
        self.assertIn('separate PDF', body)
        self.assertIn('reimbursement', body)

    def test_no_duplicate_send(self):
        with tempfile.TemporaryDirectory() as tmp:
            send = Mock()
            now = dt.datetime(2026, 10, 31, 18, tzinfo=TZ)
            job.run(now, Path(tmp), send)
            job.run(now, Path(tmp), send)
            self.assertEqual(send.call_count, 1)

    def test_ambiguous_failure_does_not_resend(self):
        with tempfile.TemporaryDirectory() as tmp:
            send = Mock(side_effect=TimeoutError('uncertain outcome'))
            now = dt.datetime(2026, 10, 31, 18, tzinfo=TZ)
            with self.assertRaises(TimeoutError):
                job.run(now, Path(tmp), send)
            send.reset_mock(side_effect=True)
            with self.assertRaises(RuntimeError):
                job.run(now, Path(tmp), send)
            send.assert_not_called()

    def test_dry_run_never_sends_or_marks_sent(self):
        with tempfile.TemporaryDirectory() as tmp:
            send = Mock()
            job.run(dt.datetime(2026, 10, 31, 18, tzinfo=TZ), Path(tmp), send, dry_run=True)
            send.assert_not_called()
            self.assertFalse((Path(tmp) / 'state.json').exists())


if __name__ == '__main__':
    unittest.main()
