# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Mac counterpart of run-tests.ps1: runs every mac/test-*.py in turn (each
# drives its own COPY of src/PrintJob.xlsm) and reports one line per script.
# A script FAILS if it exits non-zero OR prints a FAIL line or a traceback.
#
#   python3 mac/run-tests.py                     all ported tests
#   python3 mac/run-tests.py --only paid groups  name fragments; runs matching scripts
#   python3 mac/run-tests.py --label baseline    log folder suffix
#
# Logs land in /private/tmp/PrintCostsTest/testrun-<label>.
#
# Unlike run-tests.ps1 this does not prune src/*.bak.xlsm after a green run:
# that is only safe once the whole suite is ported, and a green run here does
# not yet mean the build has been tested.
#
# Close any open PrintJob workbook first.

import argparse
import glob
import os
import re
import subprocess
import sys
import time

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from xlmac import WORK_ROOT  # noqa: E402
FAIL_RE = re.compile(r'\bFAIL\b|Traceback|Error:')

ap = argparse.ArgumentParser()
ap.add_argument('--only', nargs='+')
ap.add_argument('--label', default=time.strftime('%H%M%S'))
opts = ap.parse_args()

log_dir = os.path.join(WORK_ROOT, f'testrun-{opts.label}')
os.makedirs(log_dir, exist_ok=True)

tests = sorted(p for p in glob.glob(os.path.join(HERE, 'test-*.py'))
               if os.path.basename(p) != 'test-fixture-annexe.py')
if opts.only:
    tests = [p for p in tests if any(f in os.path.basename(p) for f in opts.only)]


def run_one(path, log):
    p = subprocess.run([sys.executable, '-I', path], capture_output=True, text=True)
    out = p.stdout + p.stderr
    with open(log, 'w', encoding='utf-8') as f:
        f.write(out)
    fails = len(FAIL_RE.findall(out))
    if p.returncode != 0 and fails == 0:
        return f'EXIT {p.returncode}'
    return f'FAIL x{fails}' if fails else 'pass'


results = []
for path in tests:
    name = os.path.basename(path)
    log = os.path.join(log_dir, name[:-3] + '.log')
    t0 = time.time()
    time.sleep(2)   # let the previous test's close settle; back-to-back opens can flake
    status = run_one(path, log)
    if status != 'pass':   # AppleEvent timing flakes pass on a clean second attempt; real failures fail twice
        time.sleep(5)
        status = run_one(path, log)
        if status == 'pass':
            status = 'pass (retry)'
    results.append(status)
    print(f'{name:<32} {status:<12} {int(time.time() - t0):4}s', flush=True)

bad = sum(1 for s in results if not s.startswith('pass'))
print(f'\n{len(results) - bad} of {len(results)} passed. Logs: {log_dir}')
sys.exit(bad)
