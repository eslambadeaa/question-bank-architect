import json, sys
sys.stdout.reconfigure(encoding='utf-8')

with open(r'e:\bank\fin_dups_spec.json', 'r', encoding='utf-8') as f:
    specs = json.load(f)

print(f'Total final specs: {len(specs)}')
by_les = {}
for s in specs:
    by_les.setdefault(s['lesson'], []).append(s)

for les, items in by_les.items():
    mcq = [x for x in items if 'اختيار' in x['qtype']]
    tf = [x for x in items if 'صح' in x['qtype']]
    print(f"Lesson: {les} ({len(items)} rows): MCQ={len(mcq)}, TF={len(tf)}")
    print(f"   MCQ rows: {[x['row'] for x in mcq]}")
    print(f"   TF rows:  {[x['row'] for x in tf]}")
