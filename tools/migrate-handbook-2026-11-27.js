// Additive Handbook migration (27 Nov 2026): adds optional page fields only.
// diff = difficulty, pre = prerequisites (titles of earlier pages), fa = Finance use,
// ex = practice exercise {steps, result}. Existing page content is not touched.
// Pages are matched by subject id + title, so page order (and saved reading progress) is unchanged.
const fs = require('fs');
const path = require('path');
process.chdir(path.join(__dirname, '..'));
const hb = JSON.parse(fs.readFileSync('handbook.json', 'utf8'));
const S = id => { const s = hb.subjects.find(x => x.id === id); if (!s) throw new Error('subject ' + id); return s; };
const P = (sid, t) => { const p = S(sid).pages.find(x => x.t === t); if (!p) throw new Error('page ' + sid + ' ' + t); return p; };
const B = 'Beginner', I = 'Intermediate', A = 'Advanced', R = 'Review';
const META = {
  xl: {
    'Excel at a glance': [B, [], 'The base for every schedule, reconciliation and report you build.'],
    'Tables & clean data': [B, ['Excel at a glance'], 'AP/AR exports and accrual schedules that grow every month.'],
    'Lookups': [B, ['Tables & clean data'], 'Bring vendor names, GL descriptions or cost centres into an export.'],
    'Conditional sums & dates': [I, ['Tables & clean data'], 'Totals by cost centre and period; fiscal months for an April–March year.'],
    'Dynamic arrays & new functions': [I, ['Lookups', 'Conditional sums & dates'], 'Lists of overdue customers or unique vendors that update by themselves.'],
    'PivotTables': [B, ['Tables & clean data'], 'Expense analysis, budget vs actual and ageing summaries.'],
    'Charts & dashboards': [I, ['PivotTables'], 'Variance bridges and monthly trend charts for management packs.'],
    'What-if & checking': [I, ['Conditional sums & dates'], 'Forecast scenarios and checking formulas before you send a file.'],
    'Interview Q&A': [R, [], ''], 'Cheat sheet': [R, [], '']
  },
  pq: {
    'Power Query at a glance': [B, ['Tables & clean data'], 'Monthly exports cleaned the same way every time.'],
    'Combine monthly files': [I, ['Power Query at a glance'], 'Twelve monthly SAP exports into one table.'],
    'Transformations': [I, ['Power Query at a glance'], 'Fill down cost centres, split columns, fix dates.'],
    'Merge: lookups & bank rec': [I, ['Transformations'], 'Bank reconciliation, accrual reversals and intercompany matching.'],
    'M, parameters & refresh': [A, ['Combine monthly files', 'Merge: lookups & bank rec'], 'One folder path parameter for the whole close workbook.'],
    'Interview Q&A': [R, [], ''], 'Cheat sheet': [R, [], '']
  },
  pbi: {
    'Power BI at a glance': [B, [], 'Shared, refreshable management reporting.'],
    'Get data': [B, ['Power BI at a glance'], 'Connect GL, AR and budget files once.'],
    'Power Query transformations': [I, ['Get data'], 'Clean exports before they reach the model.'],
    'Data model: star schema': [I, ['Power Query transformations'], 'One Sales or GL fact table with Date, Customer and Account tables.'],
    'Date table (Apr–Mar)': [I, ['Data model: star schema'], 'Fiscal months and quarters for an April–March year.'],
    'DAX basics': [I, ['Data model: star schema'], 'Total Sales, Total Cost and Margin % measures.'],
    'CALCULATE & filter context': [A, ['DAX basics'], 'Budget vs actual and one-region measures.'],
    'Iterators & table functions': [A, ['CALCULATE & filter context'], 'Weighted averages and DSO-style calculations.'],
    'Time intelligence': [A, ['Date table (Apr–Mar)', 'CALCULATE & filter context'], 'YTD, last year and month-on-month comparisons.'],
    'Visuals & formatting': [I, ['DAX basics'], 'Clear variance and trend visuals.'],
    'Interactivity': [I, ['Visuals & formatting'], 'Slicers, drill-through and bookmarks for management.'],
    'Dashboard design': [I, ['Visuals & formatting'], 'A one-page CFO dashboard.'],
    'Service, sharing & RLS': [A, ['Dashboard design'], 'Each company sees only its own numbers.'],
    'Performance & advanced': [A, ['Service, sharing & RLS'], 'Fast reports on large GL data.'],
    'Interview Q&A': [R, [], ''], 'Cheat sheet': [R, [], '']
  },
  pa: {
    'Power Automate at a glance': [B, [], 'Repeat admin work done for you.'],
    'Cloud flow or desktop flow?': [B, ['Power Automate at a glance'], 'Pick the right kind of automation for a Finance task.'],
    'Save email attachments': [B, ['Power Automate at a glance'], 'Bank advices and vendor invoices saved to SharePoint automatically.'],
    'Approvals & reminders': [I, ['Save email attachments'], 'Payment requests and journal approvals with a record.'],
    'Excel, conditions & loops': [I, ['Approvals & reminders'], 'Act on rows of an Excel Table, e.g. overdue reminders.'],
    'Errors, sharing & desktop flows': [A, ['Excel, conditions & loops'], 'Flows that keep running when you are on leave.'],
    'Interview Q&A': [R, [], ''], 'Cheat sheet': [R, [], '']
  }
};
const EX = {
  'xl|Tables & clean data': [['Open a copy of last month\'s AP export', 'Click inside the data and press Ctrl+T', 'Name the Table tblAP on the Table Design tab', 'Turn on Total Row and add one new line under the Table'], 'The new line joins the Table, the total updates, and formulas copy down by themselves.'],
  'xl|Lookups': [['Put a vendor master (code, name) on a second sheet as a Table', 'In the AP Table add a Vendor name column', 'Use XLOOKUP on the vendor code with "Not found" as the if-not-found text', 'Filter for "Not found"'], 'Every invoice shows its vendor name; the filter shows only codes missing from the master.'],
  'xl|Conditional sums & dates': [['Take an expense list with Date, Cost centre and Amount', 'Add a fiscal month column for an April–March year', 'Build a small grid of cost centres by fiscal month with SUMIFS', 'Check one cell against a filtered total'], 'A grid of totals that matches the filtered totals exactly.'],
  'xl|PivotTables': [['Format the expense list as a Table', 'Insert › PivotTable on a new sheet', 'Cost centre in Rows, month in Columns, Amount in Values', 'Add a slicer for Company'], 'One row per cost centre, months across, grand totals, filtered by the slicer.'],
  'xl|Charts & dashboards': [['List budget profit, five variance lines and actual profit', 'Insert › Waterfall chart', 'Set budget and actual as totals', 'Add data labels'], 'A bridge from budget to actual that shows which variances moved profit most.'],
  'pq|Combine monthly files': [['Save three monthly exports with the same columns in one folder', 'Data › Get Data › From Folder › Combine & Transform', 'Fix types in the sample file query', 'Close & Load, then add a fourth file and Refresh All'], 'One table with all months, and the fourth month appears after Refresh.'],
  'pq|Merge: lookups & bank rec': [['Load the bank statement and the cash book as two queries', 'Merge on amount and date with Left Anti join', 'Do the same the other way round', 'Load both results'], 'Two short lists: bank items not in the books, and book items not in the bank.'],
  'pbi|Data model: star schema': [['Load Sales, Customer, Product and Date tables', 'In Model view, relate each lookup table to Sales one-to-many', 'Hide the key columns in Sales', 'Add a slicer on Customer and a table of Product sales'], 'The slicer filters the product table correctly; no blank rows appear.'],
  'pbi|DAX basics': [['Create a measure Total Sales = SUM(Sales[Amount])', 'Create Total Cost the same way', 'Create Margin % = DIVIDE([Total Sales]-[Total Cost],[Total Sales])', 'Put all three in a table by month'], 'Three measures that change with every slicer, with no divide-by-zero errors.'],
  'pa|Save email attachments': [['Create an automated cloud flow: When a new email arrives (V3)', 'Filter on one sender and Include Attachments = Yes', 'Add Create file in a SharePoint folder for each attachment', 'Send yourself a test email'], 'The test attachment appears in the folder and the run history shows Succeeded.'],
  'pa|Approvals & reminders': [['Create a list for payment requests', 'Build a flow: when an item is created, Start and wait for an approval', 'Update the item with the outcome', 'Test with a dummy request'], 'The approver gets the request; the item shows Approved or Rejected with the approver\'s name.']
};
let n = 0;
for (const [sid, pages] of Object.entries(META)) for (const [t, [diff, pre, fa]] of Object.entries(pages)) {
  const p = P(sid, t);
  if (!p.diff) { p.diff = diff; n++; }
  if (pre.length && !p.pre) { pre.forEach(x => { if (!S(sid).pages.some(y => y.t === x) && !hb.subjects.some(s => s.pages.some(y => y.t === x))) throw new Error('pre ' + x); }); p.pre = pre; }
  if (fa && !p.fa) p.fa = fa;
}
for (const [k, [steps, result]] of Object.entries(EX)) { const [sid, t] = k.split('|'); const p = P(sid, t); if (!p.ex) p.ex = { steps, result }; }
hb.version = '2026-11-27';
fs.writeFileSync('handbook.json', JSON.stringify(hb));
console.log('handbook pages tagged:', n);
