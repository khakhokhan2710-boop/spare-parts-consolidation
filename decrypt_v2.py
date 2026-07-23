# decrypt standalone - dùng python3.9 system
import subprocess, sys

# Hardcode python3.9 system path
python = "/Library/Developer/CommandLineTools/usr/bin/python3"

script = """
import msoffcrypto, io, openpyxl

result = {}
result['sheets'] = []
result['sections'] = {}
result['stats'] = {}

# Decrypt
with open('/Users/khakhokhan/Downloads/Detail Spare parts FY2026 new (2).xlsx', 'rb') as f:
    file = msoffcrypto.OfficeFile(f)
    file.load_key(password='sp')
    buf = io.BytesIO()
    file.decrypt(buf)
    buf.seek(0)

# Save decrypted
with open('/tmp/detail_v2_decrypted.xlsx', 'wb') as out:
    out.write(buf.read())
result['decrypt'] = 'OK'

# Analyze with openpyxl
wb = openpyxl.load_workbook('/tmp/detail_v2_decrypted.xlsx', data_only=True)

for sn in wb.sheetnames:
    ws = wb[sn]
    info = {
        'name': sn,
        'rows': ws.max_row,
        'cols': ws.max_column,
    }
    result['sheets'].append(info)

    # Check sections (column B)
    section_set = set()
    for r in range(12, min(ws.max_row+1, 50)):
        s = str(ws.cell(row=r, column=2).value or '').strip()
        if s and s not in ('Section', ''):
            section_set.add(s)
    
    result['sections'][sn] = list(section_set)[:10]  # first 10

wb.close()
result['stats']['total_sheets'] = len(wb.sheetnames) if wb else len(result['sheets'])

import json
print(json.dumps(result, default=str))
"""

result = subprocess.run([python, '-c', script], capture_output=True, text=True, timeout=60)
print("STDOUT:", result.stdout)
if result.stderr:
    print("STDERR:", result.stderr[-1000:])
