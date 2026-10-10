# -*- coding: utf-8 -*-
import openpyxl
import json
import sys

sys.stdout.reconfigure(encoding='utf-8')

# Read med_dups_spec.json
with open(r'e:\bank\med_dups_spec.json', 'r', encoding='utf-8') as f:
    specs = json.load(f)

print(f"Loaded {len(specs)} Medium duplicate specs.")
