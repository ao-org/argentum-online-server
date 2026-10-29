"""Gate a real HousingService claim for the native VB6 disposable-DB suite.

Not a deployment tool. Requires the explicitly marked synthetic fixture schema.
"""
import argparse
import json
import os
from pathlib import Path
import sqlite3
import sys
import time

sys.path.insert(0, os.environ["HOUSING_TEST_PYMMO_ROOT"])
from serverhelper.HousingService import HousingService, HousingConfig


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('database')
    parser.add_argument('account', type=int)
    parser.add_argument('house', type=int)
    parser.add_argument('gate')
    args = parser.parse_args()
    path = Path(args.database).resolve()
    with sqlite3.connect(f'{path.as_uri()}?mode=ro', uri=True) as check:
        assert check.execute('SELECT marker FROM housing_test_fixture').fetchone() == ('disposable-native-housing',)
    gate = Path(args.gate)

    class Auth:
        db_module = sqlite3

        def get_connection(self):
            conn = sqlite3.connect(f'{path.as_uri()}?mode=rw', uri=True, timeout=1)
            return conn, conn.cursor()

    class GatedService(HousingService):
        @staticmethod
        def _claim(conn, aid, house):
            result = HousingService._claim(conn, aid, house)
            gate.with_suffix('.ready').write_text('transaction holds write lock')
            deadline = time.monotonic() + 10
            while not gate.with_suffix('.release').exists():
                if time.monotonic() >= deadline:
                    raise RuntimeError('native test did not release gate')
                time.sleep(.005)  # bounded gate polling, never scheduling the race by a sleep
            return result

    result = GatedService(Auth(), HousingConfig(True)).request({
        'version': 4, 'operation': 'claim', 'token': str(args.account) * 64,
        'house_id': args.house})
    gate.with_suffix('.done').write_text(json.dumps({'code': int(result['code'])}))


if __name__ == '__main__':
    main()
