"""Run all or selected backend tests and report the slowest cases."""
import argparse
import os
from pathlib import Path
import sys
import time
import unittest

ROOT = Path(__file__).resolve().parents[1]


class TimedResult(unittest.TextTestResult):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.durations = []

    def startTest(self, test):
        self.started = time.perf_counter()
        super().startTest(test)

    def stopTest(self, test):
        self.durations.append((time.perf_counter() - self.started, test.id()))
        super().stopTest(test)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('tests', nargs='*', help='Optional unittest modules/classes/methods, e.g. test_account_email')
    parser.add_argument('--mysql', action='store_true', help='Enable owned loopback MySQL TEMPORARY-table tests')
    parser.add_argument('--durations', type=int, default=5, help='Number of slow cases to report (default: 5)')
    args = parser.parse_args()
    if args.mysql:
        os.environ['PRESENCE_MYSQL_TEMP_TEST'] = '1'
    sys.path.insert(0, str(ROOT))
    loader = unittest.TestLoader()
    suite = (loader.loadTestsFromNames(args.tests) if args.tests else
             loader.discover(str(ROOT), pattern='test*.py', top_level_dir=str(ROOT)))
    result = unittest.TextTestRunner(resultclass=TimedResult).run(suite)
    if args.durations > 0:
        print('\nSlowest cases (setup and cleanup included):')
        for seconds, name in sorted(result.durations, reverse=True)[:args.durations]:
            print(f'{seconds:7.3f}s {name}')
    return 0 if result.wasSuccessful() else 1


if __name__ == '__main__':
    raise SystemExit(main())
