"""Annual totals through real synthetic leave approvals and cancellation transitions."""
from contextlib import closing
import unittest
from test_leave_attendance import LeaveAttendanceTests
from test_person_names import NameCursor
from backend.leave_totals import annual_leave_totals


class LeaveTotalsTests(unittest.TestCase):
    setUp = LeaveAttendanceTests.setUp
    connect = LeaveAttendanceTests.connect
    actor_as = LeaveAttendanceTests.actor_as
    create = LeaveAttendanceTests.create
    decide = LeaveAttendanceTests.decide
    config = LeaveAttendanceTests.config
    save = LeaveAttendanceTests.save
    sheet = LeaveAttendanceTests.sheet
    request_cancel = LeaveAttendanceTests.request_cancel

    def totals(self, year=2026):
        with closing(self.connect()) as db:
            rows = [dict(row) for row in db.execute('SELECT * FROM leave_requests WHERE employee_username=?',('member',))]
            return annual_leave_totals(NameCursor(db), 'member', year,
                [r for r in rows if r['start_date'][:4]<=str(year)<=r['end_date'][:4]])

    def test_half_day_and_cancellation_status_use_real_approval_cells(self):
        identifier = self.create(start_slot='afternoon',end_slot='afternoon')
        self.assertEqual(self.totals()['approved_days'],0)
        self.assertEqual(self.decide(identifier).status_code,200)
        self.assertEqual(self.totals()['approved_days'],.5)
        self.request_cancel(identifier)
        self.assertEqual(self.totals()['approved_days'],.5)
        self.assertEqual(self.decide(identifier, cancel=True).status_code,200)
        self.assertEqual(self.totals()['approved_days'],0)

    def test_duplicate_workspaces_overlaps_and_archiving_do_not_inflate_or_erase_days(self):
        self.sheet(owner='boss');self.sheet(owner='admin')
        first = self.create()
        self.assertEqual(self.decide(first).status_code,200)
        second = self.create(start_slot='afternoon',end_slot='afternoon')
        self.assertEqual(self.decide(second).status_code,200)
        with closing(self.connect()) as db,db:
            db.execute('UPDATE leave_requests SET is_archived=1 WHERE id=?',(first,))
        self.assertEqual(self.totals()['approved_days'],1)
        self.assertFalse(self.totals()['incomplete'])
        self.assertEqual(self.totals()['months'][9]['approved_days'],1)

    def test_cross_month_skips_weekends_and_holidays(self):
        identifier=self.create(end_date='2026-11-03',start_slot='afternoon',end_slot='morning')
        self.assertEqual(self.decide(identifier).status_code,200)
        result=self.totals()
        self.assertEqual(result['approved_days'],2)
        self.assertEqual(result['months'][9]['approved_days'],.5)
        self.assertEqual(result['months'][10]['approved_days'],1.5)

    def test_cross_year_keeps_boundary_half_days(self):
        identifier=self.create(start_date='2026-12-31',end_date='2027-01-04',start_slot='afternoon',end_slot='morning')
        self.assertEqual(self.decide(identifier).status_code,200)
        self.assertEqual(self.totals(2026)['approved_days'],.5)
        self.assertEqual(self.totals(2027)['approved_days'],.5)

    def test_legacy_approval_is_explicitly_incomplete_instead_of_guessing_duration(self):
        identifier=self.create()
        with closing(self.connect()) as db,db:
            db.execute("UPDATE leave_requests SET status='approved' WHERE id=?",(identifier,))
        result=self.totals()
        self.assertEqual(result['approved_days'],0)
        self.assertEqual(result['unverified_requests'],1)
        self.assertTrue(result['months'][9]['incomplete'])

    def test_corrupt_snapshot_never_invents_duration(self):
        identifier=self.create()
        self.assertEqual(self.decide(identifier).status_code,200)
        with closing(self.connect()) as db,db:
            db.execute("UPDATE leave_attendance_cells SET after_state='invalid'")
        result=self.totals()
        self.assertEqual(result['approved_days'],0)
        self.assertTrue(result['incomplete'])

    def test_all_rest_approved_request_is_known_zero(self):
        identifier=self.create(start_date='2026-10-31',end_date='2026-11-01')
        self.assertEqual(self.decide(identifier).status_code,200)
        self.assertEqual(self.totals()['approved_days'],0)
        self.assertFalse(self.totals()['incomplete'])
