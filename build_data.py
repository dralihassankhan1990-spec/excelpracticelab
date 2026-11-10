"""Build practice-web/data/*.json from the workbook + fictional extra datasets.
Run: ~/workspace/excel-training-venv/bin/python practice-web/build_data.py
Source of truth: output/Operations_Analyst_IT_Project_Excel_Workbook.xlsx + units.py
"""
import openpyxl, json, datetime, sys, random, os

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, ROOT)
from units import UNITS

wb = openpyxl.load_workbook(os.path.join(ROOT, 'output/Operations_Analyst_IT_Project_Excel_Workbook.xlsx'), data_only=True)
ws = wb['1_IT_PM_DATA']
headers = [ws.cell(row=3, column=c).value for c in range(1, 17)]
rows = []
for r in range(4, 124):
    row = []
    for c in range(1, 17):
        v = ws.cell(row=r, column=c).value
        if isinstance(v, datetime.datetime):
            v = v.strftime('%Y-%m-%d')
        row.append(v)
    rows.append(row)
for row in rows:  # Variance = Actual - Planned (live formula in the workbook)
    if row[8] is not None and row[9] is not None:
        row[10] = round(row[9] - row[8], 1)

demo_index = {1: 1, 2: 2, 3: 3, 4: 5, 5: 6, 6: 4}
for u in UNITS:
    demo_index.setdefault(u['n'], u['n'])
formulas = [{
    'n': u['n'], 'demo': demo_index[u['n']], 'name': u['name'], 'cat': u['cat'],
    'syntax': u['syntax'], 'what': u['what'], 'read': u['read'],
    'example': u['example'], 'data': u['data'], 'result': u['result'],
    'error': u['error'], 'fix': u['fix'], 'benefit': u['benefit'],
} for u in UNITS]
with open(os.path.join(HERE, 'data/formulas.json'), 'w') as f:
    json.dump(formulas, f, indent=1)

random.seed(42)
owners = ['Ayesha Khan', 'Usman Tariq', 'Bilal Ahmed', 'Fatima Zahra',
          'Hassan Raza', 'Sana Malik', 'Omar Farooq', 'Zainab Ali']
projects = ['Customer Portal Revamp', 'HRMS Upgrade', 'Inventory API', 'Service Desk Migration']
modules = ['Login & SSO', 'Leave', 'Reports', 'Payments', 'Dashboard', 'Tickets', 'Stock']
def d(m, day):
    return f'2025-{m:02d}-{day:02d}'

# --- Bug Tracker: same header style as the course dataset so exercises transfer ---
bugs_h = ['BugID', 'Project', 'Module', 'Sprint', 'Status', 'Priority', 'Severity',
          'ReportedBy', 'AssignedTo', 'StartDate', 'DueDate', 'FinishDate',
          'StoryPoints', 'PlannedHrs', 'ActualHrs', 'VarianceHrs', '%Complete']
bugs = []
for i in range(1, 41):
    sev = random.choice(['Critical', 'High', 'Medium', 'Low'])
    status = random.choice(['Open', 'In Progress', 'Resolved', 'Closed'])
    sp = random.randint(1, 5)
    ph = round(sp * 2 + random.uniform(0, 3), 1)
    ah = round(ph + random.uniform(-2, 4), 1) if status in ('Resolved', 'Closed') else None
    pct = 1 if status in ('Resolved', 'Closed') else (0.5 if status == 'In Progress' else 0.1)
    bugs.append([f'BUG-{i:03d}', random.choice(projects), random.choice(modules),
                 random.randint(1, 6), status, sev, sev,
                 random.choice(owners), random.choice(owners),
                 d(random.randint(1, 5), random.randint(1, 28)),
                 d(6, random.randint(1, 30)),
                 d(6, random.randint(1, 30)) if status in ('Resolved', 'Closed') else '',
                 sp, ph, ah, round(ah - ph, 1) if ah is not None else None, pct])

# --- Sprint Tasks ---
sprint_h = ['TaskID', 'Project', 'Module', 'Sprint', 'Owner', 'Status', 'Priority',
            'PlannedHrs', 'ActualHrs', 'VarianceHrs', 'StartDate', 'DueDate',
            'FinishDate', 'StoryPoints', '%Complete']
sprint = []
for i in range(1, 37):
    sp = random.randint(1, 8)
    ph = round(sp * 2.5 + random.uniform(-2, 4), 1)
    status = random.choice(['Completed', 'In Progress', 'Delayed', 'Planned'])
    ah = round(ph + random.uniform(-3, 5), 1) if status in ('Completed', 'Delayed') else None
    pct = 1 if status == 'Completed' else (0.6 if status == 'In Progress' else (0.3 if status == 'Delayed' else 0))
    sprint.append([f'SP-{i:03d}', random.choice(projects), random.choice(modules),
                   random.randint(1, 6), random.choice(owners), status,
                   random.choice(['Critical', 'High', 'Medium', 'Low']),
                   ph, ah, round(ah - ph, 1) if ah is not None else None,
                   d(random.randint(1, 6), random.randint(1, 28)),
                   d(6, random.randint(1, 30)),
                   d(6, random.randint(1, 30)) if status == 'Completed' else '',
                   sp, pct])

# --- Project Budget & Costs ---
budget_h = ['CostID', 'Project', 'Category', 'Sprint', 'Owner', 'Status', 'Priority',
            'Budgeted', 'Spent', 'Vendor', 'StartDate', 'DueDate', 'FinishDate', '%Complete']
budget = []
for i in range(1, 31):
    b = random.choice([5000, 8000, 12000, 15000, 20000, 25000])
    spent = round(b * random.uniform(0.4, 1.15))
    status = random.choice(['Approved', 'Pending', 'Paid'])
    due = d(random.randint(1, 6), random.randint(1, 28))
    budget.append([f'COST-{i:03d}', random.choice(projects),
                   random.choice(['Licenses', 'Consulting', 'Hardware', 'Training', 'Cloud']),
                   random.randint(1, 6), random.choice(owners), status,
                   random.choice(['Critical', 'High', 'Medium', 'Low']),
                   b, spent, random.choice(['TechSupplies Co', 'CloudServe', 'TrainPro', 'SoftVend']),
                   d(random.randint(1, 5), random.randint(1, 28)), due,
                   due if status == 'Paid' else '',
                   round(min(1, spent / b), 2)])

datasets = [
    {'id': 'it-pm', 'title': 'IT Project Data — 120 Work Items',
     'desc': 'The course dataset: 4 projects, Jan–Jun 2025. The same data as the Excel workbook.',
     'headers': headers, 'rows': rows},
    {'id': 'bugs', 'title': 'Bug Tracker — 40 Bugs',
     'desc': 'Fictional bug list for the same 4 projects. Practice COUNTIFS, FILTER, lookups and dates.',
     'headers': bugs_h, 'rows': bugs},
    {'id': 'sprint', 'title': 'Sprint Tasks — 36 Tasks',
     'desc': 'Sprint backlog with story points and hours. Practice SUMIFS, AVERAGEIFS and ranking.',
     'headers': sprint_h, 'rows': sprint},
    {'id': 'budget', 'title': 'Project Budget & Costs — 30 Items',
     'desc': 'Budget vs spent by project and category. Practice SUMIF, SUMIFS and variance.',
     'headers': budget_h, 'rows': budget},
]
with open(os.path.join(HERE, 'data/datasets.json'), 'w') as f:
    json.dump(datasets, f, indent=1)
print('formulas:', len(formulas))
print('datasets:', [(x['id'], len(x['rows']), len(x['headers'])) for x in datasets])
