import unittest
from datetime import date
from contract_periods import monthly_exceptions


class MonthlyExceptionsTests(unittest.TestCase):
    def test_legacy_sheet_exports_selected_month_only(self):
        entries = [{'date':'2026-09-01','matin':True,'aprem':True},
                   {'date':'2026-10-01','matin':True,'aprem':True},
                   {'date':'2025-09-01','matin':True,'aprem':True}]
        employee = {key:list(entries) for key in ('vacances','absences','arret')}
        result = monthly_exceptions(employee,9,2026)
        for key in employee:
            self.assertEqual(result[key],[entries[0]])
            self.assertEqual(len(employee[key]),3)

    def test_dates_and_contract_limits(self):
        employee={'contract_id':1,'sheet_period_start':date(2026,9,15),'sheet_period_end':'2026-09-30',
                  'vacances':[{'date':date(2026,9,1)},{'date':date(2026,9,15)},{'date':None}]}
        self.assertEqual(monthly_exceptions(employee,9,2026)['vacances'],[{'date':date(2026,9,15)}])
