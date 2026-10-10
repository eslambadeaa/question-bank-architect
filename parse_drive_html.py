import re
import sys

sys.stdout.reconfigure(encoding='utf-8')

with open(r'C:\Users\MaximuM-Tech\Downloads\avenger_bank_downloaded.xlsx', 'r', encoding='utf-8', errors='ignore') as f:
    html = f.read()

for m in re.finditer(r'<title>(.*?)</title>', html, re.I):
    print('Title:', m.group(1))

if 'accounts.google.com' in html:
    print('Sign-in / Authentication required!')
else:
    print('No login required, checking form/links...')
    for m in re.finditer(r'confirm=([0-9A-Za-z_]+)', html):
        print('Confirm token:', m.group(1))
