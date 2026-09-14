#!/usr/bin/env python3
"""plans.json → plans.js（ページが <script src> で読む形にする）"""
import json, sys
src = sys.argv[1] if len(sys.argv) > 1 else 'plans.json'
dst = sys.argv[2] if len(sys.argv) > 2 else 'plans.js'
plans = json.load(open(src))
open(dst, 'w').write('window.PLANS = ' + json.dumps(plans, ensure_ascii=False) + ';\n')
print(f'{len(plans)} plans -> {dst}')
