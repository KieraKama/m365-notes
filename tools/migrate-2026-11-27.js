// One-time, additive migration (27 Nov 2026). Adds new optional fields and new top-level keys.
// It never deletes or rewrites existing items; it only adds fields, and changes the few values
// listed in CHANGED (each backed by an official Microsoft page and logged in "changes").
// Safe to run twice: every addition checks first.
const fs = require('fs');
const path = require('path');
process.chdir(path.join(__dirname, '..'));
const T = '2026-11-27';
const d = JSON.parse(fs.readFileSync('data.json', 'utf8'));
const card = id => { const f = d.features.find(x => x.id === id); if (!f) throw new Error('missing card ' + id); return f; };
const byName = n => { if (!d.features.some(f => f.name === n)) throw new Error('missing card name ' + n); return n; };
const addSrc = (list, s) => { if (!list.some(x => x[1] === s[1])) list.push(s); };

// ---------- sources read for this update ----------
const SRC = {
  dyn: ['Dynamic array formulas in non-dynamic aware Excel (Microsoft Support)', 'https://support.microsoft.com/en-us/office/dynamic-array-formulas-in-non-dynamic-aware-excel-696e164e-306b-4282-ae9d-aa88f5502fa2'],
  xlookup: ['XLOOKUP function (Microsoft Support)', 'https://support.microsoft.com/en-us/excel/functions/xlookup-function'],
  let: ['LET function (Microsoft Support)', 'https://support.microsoft.com/en-us/excel/functions/let-function'],
  lambda: ['LAMBDA function (Microsoft Support)', 'https://support.microsoft.com/en-us/excel/functions/lambda-function'],
  vstack: ['VSTACK function (Microsoft Support)', 'https://support.microsoft.com/en-us/excel/functions/vstack-function'],
  hstack: ['HSTACK function (Microsoft Support)', 'https://support.microsoft.com/en-us/office/hstack-function-98c4ab76-10fe-4b4f-8d5f-af1c125fe8c2'],
  scripts: ['Platform limits, requirements, and error messages for Office Scripts (Microsoft Learn)', 'https://learn.microsoft.com/en-us/office/dev/scripts/testing/platform-limits'],
  python: ['Python in Excel availability (Microsoft Support)', 'https://support.microsoft.com/en-us/excel/python/python-in-excel-availability'],
  pivot: ['Where is Power Pivot? (Microsoft Support)', 'https://support.microsoft.com/en-US/Excel/where-is-power-pivot'],
  pqpp: ['Learn to use Power Query and Power Pivot in Excel (Microsoft Support)', 'https://support.microsoft.com/en-gb/office/learn-to-use-power-query-and-power-pivot-in-excel-42d895c2-d1d7-41d0-88da-d1ed7ecc102d'],
  vbaweb: ['Work with VBA macros in Excel for the web (Microsoft Support)', 'https://support.microsoft.com/en-us/excel/work-with-vba-macros-in-excel-for-the-web'],
  pqweb: ['Use Power Query in Excel for the Web (Microsoft Support)', 'https://support.microsoft.com/en-us/office/use-power-query-in-excel-for-the-web-02652946-de70-48d8-8a34-3db96998ff5a'],
  pqmac: ['Import and shape data in Excel for Mac (Power Query) (Microsoft Support)', 'https://support.microsoft.com/en-us/office/import-and-shape-data-in-excel-for-mac-power-query-7b2f337d-e7d2-4fdf-bf00-3dfbb1e5e9c5'],
  morph: ['Use the Morph transition in PowerPoint (Microsoft Support)', 'https://support.microsoft.com/en-us/powerpoint/training/use-the-morph-transition-in-powerpoint'],
  zoom: ['Use zoom for PowerPoint to bring your presentation to life (Microsoft Support)', 'https://support.microsoft.com/en-us/office/use-zoom-for-powerpoint-to-bring-your-presentation-to-life-9d6c58cd-2125-4d29-86b1-0097c7dc47d7'],
  pbiUser: ['Licenses and subscriptions for business users (Microsoft Learn)', 'https://learn.microsoft.com/en-us/power-bi/fundamentals/end-user-license'],
  ppu: ['Power BI Premium Per User (Microsoft Learn)', 'https://learn.microsoft.com/en-us/fabric/enterprise/powerbi/service-premium-per-user-faq'],
  fabricLic: ['Understand Microsoft Fabric licenses and capacity (Microsoft Learn)', 'https://learn.microsoft.com/en-us/fabric/enterprise/licenses'],
  paFaq: ['Power Automate licensing FAQ (Microsoft Learn)', 'https://learn.microsoft.com/en-us/power-platform/admin/power-automate-licensing/faqs'],
  paTypes: ['Types of Power Automate licenses (Microsoft Learn)', 'https://learn.microsoft.com/en-us/power-platform/admin/power-automate-licensing/types'],
  ppFaq: ['Power Platform licensing FAQs (Microsoft Learn)', 'https://learn.microsoft.com/en-us/power-platform/admin/powerapps-flow-licensing-faq'],
  ppOverview: ['Licensing overview for Microsoft Power Platform (Microsoft Learn)', 'https://learn.microsoft.com/en-us/power-platform/admin/pricing-billing-skus'],
  copReq: ['Microsoft Copilot requirements (Microsoft Learn)', 'https://learn.microsoft.com/en-us/copilot/plan/microsoft-copilot-requirements'],
  teamsPrem: ['Microsoft Teams Premium licensing (Microsoft Learn)', 'https://learn.microsoft.com/en-us/microsoftteams/teams-add-on-licensing/licensing-enhance-teams'],
  teamsEvents: ['Teams events licensing overview (Microsoft Learn)', 'https://learn.microsoft.com/en-us/microsoftteams/events-features'],
  plannerAdmin: ['Microsoft Planner for admins (Microsoft Learn)', 'https://learn.microsoft.com/en-gb/planner/planner-for-admins'],
  plannerFaq: ['Frequently asked questions about Microsoft Planner (Microsoft Support)', 'https://support.microsoft.com/en-us/office/frequently-asked-questions-about-microsoft-planner-d1a2d4e6-a4d7-408c-a48a-31caaa267de5'],
  copLic: ['What Copilot license do I have (Microsoft Support)', 'https://support.microsoft.com/en-us/microsoft-365-copilot/what-copilot-license-do-i-have'],
  xlCopFaq: ['What is Copilot in Excel? FAQ (Microsoft Support)', 'https://support.microsoft.com/excel/copilot/frequently-asked-questions-about-copilot-in-excel'],
  pbiCop: ['Overview of Copilot for Power BI (Microsoft Learn)', 'https://learn.microsoft.com/en-us/power-bi/create-reports/copilot-introduction']
};

// ---------- 1. Finance search synonyms (used on top of the existing search) ----------
d.synonyms = [
  ['ar', 'accounts receivable', 'receivable', 'receivables', 'debtors', 'debtor', 'customer ageing', 'customer aging', 'ageing', 'aging', 'overdue'],
  ['ap', 'accounts payable', 'payable', 'payables', 'creditors', 'creditor', 'supplier', 'suppliers', 'vendor', 'vendors'],
  ['p&l', 'pnl', 'p and l', 'profit and loss', 'income statement'],
  ['bgt', 'budget', 'budgets'],
  ['fct', 'fc', 'forecast', 'forecasts', 'forecasting'],
  ['act', 'actual', 'actuals'],
  ['fx', 'forex', 'foreign exchange', 'exchange rate', 'currency'],
  ['gl', 'general ledger', 'ledger'],
  ['cogs', 'cost of goods sold', 'cost of sales'],
  ['month-end', 'month end', 'closing', 'monthly closing', 'period end'],
  ['sap export', 'sap extract', 'sap report', 'sap download', 'sap'],
  ['variance', 'variances', 'actual vs budget', 'budget vs actual', 'budget vs forecast', 'actual vs forecast', 'bva'],
  ['capex', 'capital expenditure'],
  ['fixed asset', 'fixed assets', 'asset register', 'fa register'],
  ['pbc', 'prepared by client', 'audit request', 'audit requests'],
  ['ic', 'intercompany', 'inter-company', 'inter company'],
  ['bank rec', 'bank reconciliation', 'bank recon'],
  ['rec', 'recon', 'reconciliation', 'reconcile'],
  ['dso', 'days sales outstanding'],
  ['accrual', 'accruals', 'accrued'],
  ['inventory', 'stock', 'slow-moving', 'slow moving'],
  ['pivot', 'pivottable', 'pivot table'],
  ['cash flow', 'cashflow']
];

// ---------- 2. Recommended for Finance & Accounts (references existing cards only) ----------
d.finance = {
  intro: 'A short list of existing cards that matter most for Finance & Accounts work. Start at the top of each group.',
  groups: [
    ['Excel foundation', ['FEAT-EXCEL-001', 'FEAT-EXCEL-008', 'FEAT-EXCEL-009', 'FEAT-EXCEL-013', 'FEAT-EXCEL-011', 'FEAT-EXCEL-012', 'FEAT-EXCEL-020', 'FEAT-EXCEL-007', 'FEAT-EXCEL-026', 'FEAT-EXCEL-031']],
    ['Data preparation', ['FEAT-PQ-001', 'FEAT-PQ-006', 'FEAT-PQ-005', 'FEAT-PQ-003', 'FEAT-PQ-004']],
    ['Analysis & reporting', ['FEAT-EXCEL-025', 'FEAT-PBI-001', 'FEAT-PBI-002', 'FEAT-PBI-003', 'FEAT-PBI-014']],
    ['Automation & collaboration', ['FEAT-OUTLOOK-001', 'FEAT-FLOW-002', 'FEAT-FLOW-004', 'FEAT-SP-001', 'FEAT-SP-002', 'FEAT-LISTS-001', 'FEAT-PLANNER-001']],
    ['Sharing safely & AI', ['FEAT-EXCEL-040', 'FEAT-COPILOT-001']]
  ]
};
d.finance.groups.forEach(g => g[1].forEach(card));

// ---------- 3. Learning paths (reference existing cards and Handbook pages) ----------
d.paths = [
  { id: 'L1', name: 'Level 1: Excel foundation', goal: 'Build clean, checkable workbooks for month-end lists and schedules.', steps: [
    { t: 'Excel Tables', card: 'FEAT-EXCEL-001', hb: 'xl|Tables & clean data' },
    { t: 'Structured references (Table names in formulas)', card: 'FEAT-EXCEL-001', note: 'Step 4 of the Tables card: =SUM(tblAP[Amount])' },
    { t: 'IF and logical formulas', card: 'FEAT-EXCEL-013' },
    { t: 'SUMIFS', card: 'FEAT-EXCEL-008', hb: 'xl|Conditional sums & dates' },
    { t: 'XLOOKUP', card: 'FEAT-EXCEL-009', hb: 'xl|Lookups' },
    { t: 'Date formulas (fiscal year)', card: 'FEAT-EXCEL-011', hb: 'xl|Conditional sums & dates' },
    { t: 'PivotTables', card: 'FEAT-EXCEL-020', hb: 'xl|PivotTables' },
    { t: 'Charts', card: 'FEAT-EXCEL-026', hb: 'xl|Charts & dashboards' } ] },
  { id: 'L2', name: 'Level 2: Data preparation', goal: 'Turn messy exports into one clean table you refresh instead of rebuild.', steps: [
    { t: 'Power Query fundamentals', hb: 'pq|Power Query at a glance' },
    { t: 'Import data', hb: 'pbi|Get data', note: 'Same Get Data window as Excel' },
    { t: 'Clean data', card: 'FEAT-PQ-006', hb: 'pq|Transformations' },
    { t: 'Change data types (with Locale)', card: 'FEAT-PQ-005' },
    { t: 'Merge', card: 'FEAT-PQ-003', hb: 'pq|Merge: lookups & bank rec' },
    { t: 'Append', card: 'FEAT-PQ-004' },
    { t: 'From Folder', card: 'FEAT-PQ-001', hb: 'pq|Combine monthly files' },
    { t: 'Refresh', card: 'FEAT-PQ-012', hb: 'pq|M, parameters & refresh' },
    { t: 'Finding errors (Column quality & profile)', card: 'FEAT-PQ-010' } ] },
  { id: 'L3', name: 'Level 3: Analysis & reporting', goal: 'Model the data once and report it interactively.', steps: [
    { t: 'Data Model', card: 'FEAT-EXCEL-025' },
    { t: 'Relationships (star schema)', card: 'FEAT-PBI-001', hb: 'pbi|Data model: star schema' },
    { t: 'Power Pivot', card: 'FEAT-EXCEL-025' },
    { t: 'Power BI fundamentals', hb: 'pbi|Power BI at a glance' },
    { t: 'Visualisations', hb: 'pbi|Visuals & formatting' },
    { t: 'Measures', card: 'FEAT-PBI-003' },
    { t: 'DAX fundamentals', hb: 'pbi|DAX basics' },
    { t: 'Interactive reporting', card: 'FEAT-PBI-008', hb: 'pbi|Interactivity' } ] },
  { id: 'L4', name: 'Level 4: Automation', goal: 'Automate the repeat work, with owners and controls in place.', steps: [
    { t: 'Power Automate', card: 'FEAT-FLOW-001', hb: 'pa|Power Automate at a glance' },
    { t: 'SharePoint / Lists workflows', card: 'FEAT-FLOW-005' },
    { t: 'Approvals', card: 'FEAT-FLOW-004', hb: 'pa|Approvals & reminders' },
    { t: 'Scheduled processes', card: 'FEAT-FLOW-003' },
    { t: 'Office Scripts (where appropriate)', card: 'FEAT-EXCEL-043' },
    { t: 'Controlled automation (error handling, co-owners)', card: 'FEAT-FLOW-010', hb: 'pa|Errors, sharing & desktop flows' } ] },
  { id: 'L5', name: 'Level 5: AI productivity', goal: 'Use Copilot well and safely for Finance work.', steps: [
    { t: 'Microsoft Copilot: which one do I have?', card: 'FEAT-COPILOT-001', hb: 'cp|Copilot at a glance' },
    { t: 'Good prompting', hb: 'cp|Write a good prompt' },
    { t: 'Excel Copilot', card: 'FEAT-COPILOT-019', hb: 'cpp|Copilot in Excel' },
    { t: 'Word Copilot', card: 'FEAT-COPILOT-018', hb: 'cpp|Copilot in Word' },
    { t: 'PowerPoint Copilot', card: 'FEAT-COPILOT-020', hb: 'cpp|Copilot in PowerPoint' },
    { t: 'Outlook Copilot', card: 'FEAT-COPILOT-006' },
    { t: 'Teams Copilot', card: 'FEAT-COPILOT-013' },
    { t: 'Finance use cases', hb: 'cp|Everyday jobs with Basic' },
    { t: 'Data security and responsible use', hb: 'cp|Use it safely and check the answer', link: 'safety' } ] }
];
d.paths.forEach(p => p.steps.forEach(s => { if (s.card) card(s.card); }));

// ---------- 4. "Try it" exercises on existing cards (new optional fields tryit / result) ----------
const TRY = {
  'FEAT-EXCEL-001': ['Turn last month\'s AP export into a Table named tblAP and switch on Total Row (Table Design tab).', 'Banded rows with filter buttons, a total under the Amount column, and formulas that read tblAP[Amount].'],
  'FEAT-EXCEL-008': ['On an expense list, total one cost centre for one month with SUMIFS.', 'One number that matches the total you get by filtering the same rows.'],
  'FEAT-EXCEL-009': ['Bring the vendor name into an AP list from a vendor master, using the vendor code.', 'Every row shows a vendor name; codes missing from the master show your "not found" text instead of #N/A.'],
  'FEAT-EXCEL-013': ['Add a status column that shows "Overdue" when the due date is before today, otherwise "Current".', 'A status column you can filter and count, with no error values on blank rows.'],
  'FEAT-EXCEL-011': ['From an invoice date, work out the fiscal year and fiscal month for an April–March year.', 'Each date shows its fiscal year and period. Check a few March and April dates by hand.'],
  'FEAT-EXCEL-012': ['Round a column of FX-converted amounts to 2 decimals with ROUND and compare the total with the unrounded total.', 'You can see the rounding difference and decide where to post it.'],
  'FEAT-EXCEL-020': ['Create a PivotTable showing expenses by cost centre (rows) and month (columns).', 'One row per cost centre, months across the columns, and grand totals.'],
  'FEAT-EXCEL-007': ['Highlight AR lines more than 90 days overdue.', 'Only those rows change colour, and they update by themselves as dates move.'],
  'FEAT-EXCEL-026': ['Build a waterfall chart from budget profit to actual profit with the main variances in between.', 'A bridge: starting bar, up and down bars for each variance, ending bar.'],
  'FEAT-EXCEL-031': ['Use Evaluate Formula on a SUMIFS that gives an unexpected total.', 'You see each step of the calculation and which argument gives the wrong result.'],
  'FEAT-EXCEL-025': ['Add a sales table and a customer table to the Data Model, relate them by customer code, and build one PivotTable from both.', 'A PivotTable of sales by customer group with no lookup columns in the sales sheet.'],
  'FEAT-EXCEL-040': ['Before emailing a workbook outside the company, run Check Compatibility and Inspect Document on a copy.', 'A list of features that may break in older Excel, and hidden data you can remove.'],
  'FEAT-PQ-001': ['Put three monthly exports with the same columns in one folder and combine them with From Folder.', 'One table with every row, plus a Source.Name column showing which file each row came from.'],
  'FEAT-PQ-003': ['Merge a bank statement with the cash book, then use a Left Anti join to list bank lines with no match.', 'A short list of unmatched bank lines to investigate.'],
  'FEAT-PQ-004': ['Append the AR lists of two companies into one query.', 'One table with both companies\' rows; columns with the same name line up.'],
  'FEAT-PQ-005': ['Load a CSV whose dates are day/month/year and set the Date type using the locale the file was made in.', 'Dates read correctly (5 April stays 5 April), with no errors or swapped day and month.'],
  'FEAT-PQ-006': ['On an export with the cost centre only on the first row of each block, Fill Down, Trim and remove blank rows.', 'Every row has its cost centre and there are no blank or padded rows.'],
  'FEAT-PQ-012': ['Open Query Dependencies to see which files the workbook reads, then Refresh All.', 'A diagram from source files to queries; after refresh the loaded row counts change to the new data.'],
  'FEAT-PBI-001': ['Build a model with one Sales table and Date, Customer and Product tables related one-to-many.', 'Slicers on Customer and Date filter every sales visual correctly.'],
  'FEAT-PBI-003': ['Write Total Sales = SUM(Sales[Amount]) and a second measure that uses CALCULATE to show one region only.', 'Both measures react to slicers; the second always stays on that region.'],
  'FEAT-OUTLOOK-001': ['Make a rule that moves bank advice emails from one sender into a "Bank advices" folder.', 'New emails from that sender go straight into the folder.'],
  'FEAT-FLOW-002': ['Make a flow that saves PDF attachments from one sender into a SharePoint folder.', 'After the flow runs, each new email\'s PDF is in the folder; the run history shows Succeeded.'],
  'FEAT-FLOW-004': ['Build an approval for a payment request item in a list.', 'The approver gets the request in Teams and email, and the item status changes to Approved or Rejected.'],
  'FEAT-SP-001': ['Open the version history of a shared reconciliation file and view an earlier version.', 'A list of versions with who changed the file and when; you can open or restore one.'],
  'FEAT-LISTS-001': ['Create a list to track audit PBC requests with owner, due date and status.', 'A shared tracker the team can filter by owner and status.'],
  'FEAT-PLANNER-001': ['Create a month-end plan with buckets D+1, D+2, D+3 and a recurring task.', 'A board of close tasks; the recurring task comes back each month.']
};
for (const [id, [t, r]] of Object.entries(TRY)) { const f = card(id); if (!f.tryit) { f.tryit = t; f.result = r; f.updated = T; } }

// ---------- 5. Tool comparison for "Which tool for the job?" ----------
d.toolGuide = {
  note: 'General guidance, not a rule. Your company\'s licences, data policy and IT set-up decide what you can use.',
  cols: ['Tool', 'Best for', 'Difficulty', 'Automation', 'Collaboration', 'Licence', 'IT dependency', 'Limitations'],
  rows: [
    ['x', 'Excel', 'Quick or individual analysis, schedules, checks', 'Basic → Advanced', 'Manual; macros or Office Scripts if allowed', 'Co-authoring when saved in OneDrive/SharePoint', 'Included in Microsoft 365 plans', 'Macros may be blocked', '1,048,576 rows per sheet; manual steps are easy to get wrong'],
    ['q', 'Power Query', 'Repeatable cleaning and combining of exports', 'Intermediate', 'One click Refresh; not scheduled on its own', 'Share the workbook; sources must be reachable by others', 'Included in Excel', 'Sources on someone\'s C: drive break for others', 'Excel for the web refreshes only some sources'],
    ['b', 'Power BI', 'Interactive, recurring reports for many people', 'Intermediate → Advanced', 'Scheduled refresh in the service', 'Workspaces and apps', 'Pro or PPU to share; Free viewers only on F64+ capacity', 'Gateway for on-premises sources; export may be restricted', 'Needs a data model; learning DAX takes time'],
    ['a', 'Power Automate', 'Notifications, approvals, moving files and data', 'Intermediate', 'Fully automated, scheduled or triggered', 'Co-owners; runs in the cloud', 'Standard connectors with Microsoft 365; premium connectors need Premium', 'DLP policies decide which connectors can be combined', 'Flows stop if the owner leaves or connections expire'],
    ['l', 'Microsoft Lists', 'Structured shared tracking (requests, issues, PBC)', 'Basic', 'Rules and alerts; Power Automate', 'Everyone edits one list', 'Included in Microsoft 365 plans', 'Low', 'Not for calculations or large data'],
    ['pl', 'Planner', 'Team task management with owners and dates', 'Basic', 'Recurring tasks', 'Shared plans in Teams', 'Basic plans included; premium plans need Planner Plan 1', 'Low', 'Not a data tracker'],
    ['s', 'SharePoint', 'Controlled team storage with versions and permissions', 'Basic', 'Retention and flows', 'Shared libraries and links', 'Included in Microsoft 365 plans', 'External sharing often restricted', 'Moving or renaming folders can break links and queries'],
    ['p', 'Power Apps', 'Simple data-entry apps on top of lists', 'Intermediate → Advanced', 'With Power Automate', 'Shared app', 'Standard connectors with Microsoft 365; Dataverse or premium connectors need Power Apps Premium per user', 'Environments and DLP', 'Needs maintenance like any small system'],
    ['c', 'Copilot', 'Drafting, summarising, explaining and assisted analysis', 'Basic', 'Agents; Copilot in flows', 'Pages and notebooks', 'Copilot Chat with work account; Microsoft 365 Copilot per user for full use', 'Admins can turn features off; follow the AI policy', 'Can be wrong; you check every number']
  ],
  src: [SRC.pbiUser, SRC.fabricLic, SRC.paFaq, SRC.ppFaq, SRC.plannerAdmin, SRC.copReq, SRC.pqweb]
};

// ---------- 6. New Finance scenarios for "Which tool for the job?" ----------
const SAPNOTE = 'Microsoft side only: SAP transactions, report layouts and your SAP permissions are set by your SAP team and IT. Nothing here posts to SAP.';
const NEWTASKS = [
  { task: 'Consolidate monthly forecast files from departments', size: 'Medium', scope: 'fin', app: 'q',
    why: 'Power Query From Folder combines every department\'s file into one table you refresh each month.',
    alt: 'Copy into one Excel Table for two or three small files; Forms or Lists if you want structured input instead of files.',
    pic: { k: 'flow', steps: ['Same template for every department', 'Files saved in one SharePoint folder', 'Power Query: From Folder › Combine', 'PivotTable by department and month · Refresh'] },
    steps: ['Give every department the same Excel template (same sheet name and columns)', 'Save the returned files in one SharePoint folder', 'Data › Get Data › From Folder (or SharePoint Folder) › Combine & Transform', 'Load to a PivotTable by department and month; next month drop in the new files and Refresh All'],
    related: ['Get Data from Folder (combine files)', 'Excel Tables (Ctrl+T)', 'PivotTable'],
    options: [['One-off, two or three files', 'Excel', 'Copy into one Table; quick and visible'], ['Every month, many files', 'Power Query From Folder', 'Same steps every month with one Refresh'], ['Many people need the result', 'Power BI', 'Shared, refreshable report; sharing needs Pro (or capacity)']] },
  { task: 'Actual vs Budget vs Forecast (three-way variance)', size: 'Medium', scope: 'fin', app: 'x',
    why: 'Stack actual, budget and forecast in one Table with a Version column; one PivotTable then compares all three.',
    alt: 'Power BI with a Version column and variance measures for a recurring, shared report.',
    pic: { k: 'pivot', cols: ['Account', 'Actual', 'Budget', 'Forecast', 'Var vs Bgt', 'Var vs Fct'], rows: [['Revenue', '1,040', '1,000', '1,060', '40', '-20'], ['Freight', '118', '100', '110', '18', '8'], ['Utilities', '76', '80', '78', '-4', '-2']], hl: [1] },
    steps: ['Put the three sets in one Table: Account, Period, Version (Actual/Budget/Forecast), Amount; Append in Power Query helps', 'Insert › PivotTable: Account in Rows, Version in Columns, Amount in Values', 'Add variances with Show Values As › Difference From, or formulas beside the pivot', 'Add a Period slicer and highlight large variances'],
    related: ['Append queries', 'PivotTable', "Pivot 'Show Values As' & calculated fields", 'Slicers & Timeline'],
    options: [['One month, one entity', 'Excel PivotTable', 'Fast and easy to check'], ['Every month, several entities', 'Power Query + PivotTable or Data Model', 'Refresh instead of rebuild'], ['Management wants to slice it themselves', 'Power BI', 'Interactive and shared']] },
  { task: 'Monthly management reporting pack', size: 'Medium', scope: 'fin', app: 'pp',
    why: 'Keep the numbers in one Excel workbook and paste-link charts and tables into PowerPoint, so the pack updates with the numbers.',
    alt: 'Power BI report exported to PowerPoint or PDF; paginated reports for fixed-format packs.',
    pic: { k: 'slides', title: 'Management pack: October', bullets: ['P&L vs budget', 'Cash and working capital', 'Key variances explained'] },
    steps: ['Build each pack page as a sheet in one Excel workbook', 'Copy a chart › PowerPoint › Paste Special › Paste link', 'Next month update the workbook, then update the links in PowerPoint (File › Info › Edit Links)', 'Send the final pack as PDF'],
    related: ['Paste-link Excel charts & tables', 'Slide Master & layouts', 'Charts: waterfall, combo & sparklines'],
    options: [['Fixed pack, few readers', 'Excel + PowerPoint (paste link)', 'Familiar format, easy to print'], ['Readers want to drill down', 'Power BI', 'Interactive, one version of the numbers'], ['Final numbers going outside the team', 'PDF', 'Fixed layout that can\'t be changed by accident']] },
  { task: 'Expense analysis by cost centre and account', size: 'Small', scope: 'fin', app: 'x',
    why: 'A PivotTable on the GL expense export shows totals by cost centre, account and month in seconds.',
    alt: 'Power BI if the same analysis is needed every month by several people.',
    pic: { k: 'pivot', cols: ['Cost centre', 'Travel', 'Utilities', 'Repairs', 'Total'], rows: [['CC100', '12', '30', '4', '46'], ['CC200', '8', '22', '15', '45'], ['Total', '20', '52', '19', '91']], hl: [2] },
    steps: ['Format the export as a Table (Ctrl+T)', 'Insert › PivotTable: Cost centre in Rows, Account in Columns, Amount in Values', 'Add slicers for Month and Company', 'Double-click a number to see the lines behind it'],
    related: ['Excel Tables (Ctrl+T)', 'PivotTable', 'Slicers & Timeline'] },
  { task: 'Accrual reversal monitoring', size: 'Medium', scope: 'fin', app: 'q',
    why: 'Merge last month\'s accrual list with this month\'s GL lines on the reference; a Left Anti join lists accruals not yet reversed.',
    alt: 'XLOOKUP with a "not found" text in Excel for a short list.',
    pic: { k: 'rec', left: ['Accruals Oct', ['ACR-01 1,200', 'ACR-02 800', 'ACR-03 450']], right: ['GL Nov reversals', ['ACR-01 -1,200', 'ACR-02 -800']], out: ['Not reversed: ACR-03 450'] },
    steps: ['Keep the accrual list as a Table with a unique reference per accrual', 'Load it and the next month\'s GL export into Power Query', 'Merge on the reference with Left Anti join', 'Load the result: these accruals have no reversal yet'],
    related: ['Merge queries (joins), incl. anti join', 'Excel Tables (Ctrl+T)', 'Conditional formatting'] },
  { task: 'Accounts payable ageing and analysis', size: 'Medium', scope: 'fin', app: 'x',
    why: 'Days past due from the due date, an ageing bucket column and a PivotTable give AP ageing by vendor.',
    alt: 'Power BI for a recurring, shared ageing report.',
    pic: { k: 'pivot', cols: ['Vendor', 'Current', '1-30', '31-60', '>60'], rows: [['Vendor A', '5,000', '1,200', '0', '0'], ['Vendor B', '2,300', '0', '800', '150'], ['Total', '7,300', '1,200', '800', '150']], hl: [1] },
    steps: ['Export open AP items to Excel and format as a Table', 'Add Days past due = report date − due date', 'Add a bucket column with IFS (Current, 1-30, 31-60, >60)', 'PivotTable: Vendor in Rows, Bucket in Columns, Amount in Values'],
    related: ['Excel Tables (Ctrl+T)', 'IFERROR, IFS, SWITCH', 'PivotTable'] },
  { task: 'Inventory ageing and slow-moving stock', size: 'Medium', scope: 'fin', app: 'x',
    why: 'Days since last movement and an ageing bucket show slow-moving items; a PivotTable totals value by bucket.',
    alt: 'Power Query if the stock report needs cleaning every month; Power BI for a shared view.',
    pic: { k: 'sheet', cols: ['Material', 'Last movement', 'Days', 'Bucket', 'Value'], rows: [['M-100', '02/10', '56', '31-90', '4,200'], ['M-205', '15/03', '257', '>180', '9,800'], ['M-310', '20/11', '7', '0-30', '1,100']], hl: [1] },
    steps: ['Export stock with last movement date and value (as your SAP or ERP access allows)', 'Clean it in Power Query or as a Table', 'Add Days since last movement and a bucket column', 'PivotTable by plant or material group and bucket; highlight the oldest bucket'],
    related: ['Excel Tables (Ctrl+T)', 'Clean basics: Fill Down, Split, Trim, Replace', 'PivotTable'], sap: SAPNOTE },
  { task: 'Standard cost vs actual cost (production cost variance)', size: 'Medium', scope: 'fin', app: 'q',
    why: 'Merge the standard cost table with actual costs per material or order, then calculate variance and variance % for a PivotTable.',
    alt: 'XLOOKUP in Excel for one month; Power BI for trends across months.',
    pic: { k: 'pivot', cols: ['Product', 'Standard', 'Actual', 'Variance', 'Var %'], rows: [['P-01', '10.00', '10.60', '0.60', '6%'], ['P-02', '8.50', '8.20', '-0.30', '-4%'], ['P-03', '12.00', '13.80', '1.80', '15%']], hl: [2] },
    steps: ['Export standard costs and actual costs (as your SAP or ERP access allows)', 'In Power Query, Merge actuals with standards on material or order', 'Add Variance = Actual − Standard and Variance % = Variance / Standard', 'PivotTable by product and month; highlight variances above your threshold'],
    related: ['Merge queries (joins), incl. anti join', 'PivotTable', 'Conditional formatting'], sap: SAPNOTE },
  { task: 'Realised FX gain/loss reconciliation', size: 'Medium', scope: 'fin', app: 'q',
    why: 'Merge invoices (booking rate) with payments (payment rate) on the document number to recalculate realised FX and compare with the GL account.',
    alt: 'XLOOKUP in Excel for a handful of payments.',
    pic: { k: 'sheet', cols: ['Doc', 'USD', 'Book rate', 'Pay rate', 'Realised'], rows: [['INV-1', '1,000', '4.40', '4.45', '50'], ['INV-2', '2,500', '4.38', '4.35', '-75'], ['', '', '', 'Total', '-25']], hl: [2] },
    steps: ['Export paid foreign-currency invoices and their payments', 'In Power Query, Merge payments with invoices on the document number', 'Add Realised FX = foreign amount × (payment rate − booking rate), using your company\'s sign convention', 'Total it and compare with the realised FX GL account; investigate differences'],
    related: ['Merge queries (joins), incl. anti join', 'ROUND & precision control', 'PivotTable'] },
  { task: 'Unrealised FX revaluation check', size: 'Medium', scope: 'fin', app: 'x',
    why: 'XLOOKUP the closing rate onto open foreign-currency items and recalculate the revaluation to check what was posted.',
    alt: 'Power Query merge if the open-items list needs cleaning first.',
    pic: { k: 'sheet', cols: ['Item', 'USD', 'Book rate', 'Close rate', 'Unrealised'], rows: [['AR-1', '3,000', '4.40', '4.42', '60'], ['AP-7', '-1,500', '4.36', '4.42', '-90'], ['', '', '', 'Total', '-30']], hl: [2] },
    steps: ['Export open foreign-currency items at month-end', 'Keep closing rates in a small Table', 'XLOOKUP the closing rate by currency; Unrealised = amount × (closing − booking rate)', 'Total by account and compare with the revaluation posted; round with ROUND'],
    related: ['XLOOKUP', 'ROUND & precision control', 'Excel Tables (Ctrl+T)'] },
  { task: 'Intercompany reconciliation', size: 'Medium', scope: 'fin', app: 'q',
    why: 'Append both companies\' intercompany balances, then Merge by counterparty and currency to show differences and one-sided items.',
    alt: 'A PivotTable with company and counterparty for a quick check.',
    pic: { k: 'rec', left: ['Company A books', ['B: 120,000', 'C: 45,000']], right: ['Company B books', ['A: -120,000', 'A (in transit): -2,000']], out: ['Difference A↔B: 0 after in-transit item', 'Check C: no matching balance received'] },
    steps: ['Export IC balances from each company with counterparty, currency and amount', 'Append them in Power Query', 'Group By pair and currency, or Merge each side against the other', 'Investigate pairs that do not net to zero; use an anti join for one-sided items'],
    related: ['Append queries', 'Merge queries (joins), incl. anti join', 'Group By'] },
  { task: 'CAPEX tracking against budget', size: 'Medium', scope: 'fin', app: 'l',
    why: 'A list holds each CAPEX request with budget line, owner, approval status and amount; Excel compares committed and spent with budget.',
    alt: 'Excel Table only, for a small team; Power BI if management needs a live view.',
    pic: { k: 'list', title: 'CAPEX requests', items: [['CX-01 · Forklift', 'Approved', 1], ['CX-02 · Laptops', 'Pending', 0], ['CX-03 · Roof repair', 'Draft', 0]] },
    steps: ['Create a list with Request, Budget line, Owner, Amount, Status', 'Use an approval flow for requests above your limit', 'Export the list to Excel or connect Power Query', 'PivotTable: budget line vs approved and spent amounts'],
    related: ['Lists from templates', 'Approvals', 'PivotTable'] },
  { task: 'Fixed asset register tracking (Excel side)', size: 'Medium', scope: 'fin', app: 'x',
    why: 'An Excel Table of assets with cost, acquisition date and useful life lets you check depreciation and additions against the GL.',
    alt: 'Use the fixed-asset module of your ERP as the book of record; Excel is for checks and analysis.',
    pic: { k: 'sheet', cols: ['Asset', 'Cost', 'Start', 'Life (yrs)', 'Dep/month'], rows: [['FA-01', '36,000', '01/01/26', '5', '600'], ['FA-02', '12,000', '01/06/26', '3', '333.33']], hl: [] },
    steps: ['Export the asset list as your ERP access allows, or keep a Table', 'Add monthly depreciation using your company\'s method and rounding (ROUND)', 'Use fiscal-period formulas to show depreciation per period', 'Reconcile the totals with the GL each month'],
    related: ['Excel Tables (Ctrl+T)', 'Fiscal-year date formulas', 'ROUND & precision control'] },
  { task: 'Audit PBC request tracking', size: 'Medium', scope: 'fin', app: 'l',
    why: 'A shared list of PBC requests with owner, due date, status and file link keeps the team and the auditors on one tracker.',
    alt: 'Planner for task-style tracking; an Excel Table for a very small audit.',
    pic: { k: 'list', title: 'Audit PBC list', items: [['PBC-01 · Bank confirmations', 'Sent', 1], ['PBC-02 · FA additions', 'In progress', 0], ['PBC-03 · AR ageing', 'Not started', 0]] },
    steps: ['Create a list with Request, Owner, Due date, Status, Link to file', 'Add a rule to alert the owner when a request is assigned', 'Group the view by Status or Owner', 'Share files with auditors with Specific people links, as your IT allows'],
    related: ['Lists from templates', 'Rules & alerts', 'Sharing link types'] },
  { task: 'Cash flow analysis from bank transactions', size: 'Medium', scope: 'fin', app: 'q',
    why: 'Power Query cleans bank statement exports and Merges them with a mapping table of categories; a PivotTable shows cash in and out by week and category.',
    alt: 'Power BI for a shared, refreshable cash dashboard.',
    pic: { k: 'pivot', cols: ['Category', 'Wk 1', 'Wk 2', 'Wk 3'], rows: [['Receipts', '820', '640', '910'], ['Payroll', '-300', '0', '-300'], ['Suppliers', '-410', '-520', '-380'], ['Net', '110', '120', '230']], hl: [3] },
    steps: ['Save bank statement exports in one folder', 'Combine them with From Folder and clean the columns', 'Merge with a mapping table (description keyword → category)', 'PivotTable by week and category; check the net against the bank balance movement'],
    related: ['Get Data from Folder (combine files)', 'Merge queries (joins), incl. anti join', 'PivotTable'] },
  { task: 'Payment supporting-document workflow', size: 'Big', scope: 'fin', app: 'a',
    why: 'A list item per payment request, supporting documents in SharePoint and a Power Automate approval give one controlled, audited trail.',
    alt: 'Approvals app in Teams for low volumes; your ERP\'s own workflow if it has one.',
    pic: { k: 'flow', steps: ['Request in list + documents in SharePoint', 'Approval to the approver (Teams/email)', 'Status and approver saved on the item', 'Payment processed in ERP/bank as usual'] },
    steps: ['Create a list with Payee, Amount, Due date, Status and a link to the documents', 'Build a flow: when an item is created, start an approval', 'Write the outcome and approver back to the item', 'Process the payment in your ERP or bank with the normal controls'],
    related: ['Approvals', 'Lists from templates', 'Version history & restore'],
    gov: 'HIGH: financial approvals. The flow records approvals only; it does not post or release payments. Agree it with your finance controls and IT.' },
  { task: 'SAP extract to a refreshable Excel report', size: 'Medium', scope: 'fin', app: 'q',
    why: 'Save the SAP export with the same layout and file name in a fixed SharePoint folder; Power Query reads it and the report refreshes instead of being rebuilt.',
    alt: 'Clean one export by hand (Text to Columns) if it is a one-off.',
    pic: { k: 'flow', steps: ['SAP export, same layout each month', 'Saved to a fixed SharePoint folder', 'Power Query: clean, types, merge', 'Report sheets · Refresh All'] },
    steps: ['Agree one report layout (variant) for the export with your SAP team', 'Save each export to the same SharePoint folder with a consistent file name', 'Build the Power Query once: remove header rows, set types with Locale, merge master data', 'Each month replace the file and Refresh All; check row counts and totals'],
    related: ['Get Data from Folder (combine files)', 'Change type with Locale', 'Clean basics: Fill Down, Split, Trim, Replace', 'Parameters (file path, period)'], sap: SAPNOTE },
  { task: 'Customer ageing and overdue follow-up list', size: 'Small', scope: 'fin', app: 'x',
    why: 'Days overdue, conditional formatting and filters give the collection team a daily overdue list by customer.',
    alt: 'Power BI DSO / AR ageing report for management; Word mail merge for reminder letters.',
    pic: { k: 'sheet', cols: ['Customer', 'Invoice', 'Due', 'Days over', 'Amount'], rows: [['Customer A', 'INV-501', '01/09', '87', '4,500'], ['Customer B', 'INV-512', '15/10', '43', '1,200'], ['Customer C', 'INV-530', '20/11', '7', '900']], hl: [0] },
    steps: ['Export open AR items and format as a Table', 'Add Days overdue = report date − due date and an ageing bucket', 'Conditional formatting for more than 60 or 90 days', 'Filter by collector or customer and follow up; reminder letters with mail merge'],
    related: ['Excel Tables (Ctrl+T)', 'Conditional formatting', 'Mail Merge from Excel'] }
];
for (const t of NEWTASKS) { t.related.forEach(byName); if (!d.tasks.some(x => x.task === t.task)) d.tasks.push(t); }

// Scale guidance on existing tasks (adds fields; existing text untouched)
const OPT = {
  'Combine 12 monthly SAP exports': [['Small one-off task', 'Excel formulas or VSTACK', 'Enough for a few small files (VSTACK needs Microsoft 365 or Excel 2024)'], ['Recurring monthly process', 'Power Query From Folder', 'Drop the new file in and Refresh'], ['Shared reporting or model', 'Power BI or an IT-managed shared data source', 'One governed dataset for everyone']],
  'Bank reconciliation': [['Few lines, one account', 'Excel (XLOOKUP, filters)', 'Quick to set up'], ['Every month, many lines', 'Power Query Merge with anti joins', 'Same steps every month; unmatched items listed automatically'], ['Several accounts and people', 'Power Query plus a shared tracker (Lists) for open items', 'Clear owners and status']],
  'Budget vs actual variance': [['One-off question', 'PivotTable', 'Fast'], ['Monthly pack', 'Power Query + PivotTable or Data Model', 'Refresh instead of rebuild'], ['Many readers, drill-down', 'Power BI', 'Interactive and shared']],
  'File bank credit advices from email': [['A few emails a week', 'Outlook rule', 'No licence or IT set-up'], ['Every day, save to SharePoint', 'Power Automate (standard connectors)', 'Runs without you'], ['Read amounts from the PDFs', 'AI Builder or desktop flows', 'Needs extra licences and IT approval']],
  'Payment / journal approval': [['Occasional, small team', 'Approvals app in Teams', 'Simple, recorded'], ['Regular, with documents', 'Lists + Power Automate approval', 'Audit trail per item'], ['Posting in the ERP', 'The ERP\'s own workflow', 'Keeps postings under ERP controls']],
  'Management dashboard for several companies': [['One company, few readers', 'Excel with PivotCharts and slicers', 'No extra licence'], ['Several companies, monthly', 'Power BI', 'One model, refreshed'], ['Viewers without Pro', 'Power BI on F64+ capacity, or export to PDF', 'Free viewers need F64 or larger']],
  'Repetitive SAP GUI downloads (e.g. CN41N)': [['A few downloads', 'Manual export', 'No IT approval needed'], ['Many, repeated', 'SAP GUI scripting or Power Automate desktop', 'Needs IT/SAP approval; read-only use only'], ['Scheduled, unattended', 'An IT-managed extract', 'Unattended bots need licences and strong controls']],
  'Weekly report emailed automatically': [['Readers have Power BI', 'Power BI subscription', 'Built in'], ['Readers need a file', 'Power Automate scheduled flow', 'Sends a fixed file or link'], ['Fixed-format pages', 'Paginated reports', 'Print-ready layout']]
};
for (const [n, o] of Object.entries(OPT)) { const t = d.tasks.find(x => x.task === n); if (!t) throw new Error('task ' + n); if (!t.options) t.options = o; }
for (const n of ['Clean one SAP export for today', 'Combine 12 monthly SAP exports', 'Repetitive SAP GUI downloads (e.g. CN41N)']) { const t = d.tasks.find(x => x.task === n); if (t && !t.sap) t.sap = SAPNOTE; }

// ---------- 7. Compatibility: 4 levels, verified notes ----------
d.compatLevels = [['y', 'Works fully', '✓'], ['p', 'Opens with limitations', '~'], ['s', 'View / static only', '◐'], ['n', 'Does not work', '✗']];
d.compatNotes = [
  ['Dynamic arrays (FILTER, SORT, UNIQUE)', 'p', 'In Excel without dynamic arrays, compatible formulas may show as legacy array formulas that don\'t resize; formulas using newer functions may not calculate correctly.', SRC.dyn],
  ['XLOOKUP', 'n', 'Not available in Excel 2016 and Excel 2019. Use INDEX + MATCH when the file goes to those versions.', SRC.xlookup],
  ['LET', 'n', 'Available in Excel 2021, 2024, Microsoft 365 and the web; not in 2019 or 2016.', SRC.let],
  ['LAMBDA', 'n', 'Available in Excel 2024, Microsoft 365 and the web; not in 2021 or older.', SRC.lambda],
  ['VSTACK / HSTACK', 'n', 'Available in Excel 2024, Microsoft 365 and the web; not in 2021 or older.', SRC.vstack],
  ['Office Scripts', 'n', 'Need Excel on the web, Excel for Windows (version 2210 or later) or Excel for Mac, OneDrive for Business and a Microsoft 365 business, enterprise or education subscription. Microsoft lists only subscription licences.', SRC.scripts],
  ['Python in Excel', 's', 'Microsoft 365 business/enterprise on Windows, Mac and the web (family/personal in preview). People without it can view the workbook, but Python cells show an error when recalculated.', SRC.python],
  ['Power Query', 'p', 'Full in Excel for Windows. Excel for the web can refresh queries and import from some sources; Excel for Mac (version 16.69 or later) has the editor and refreshes many sources.', SRC.pqweb],
  ['Power Pivot', 'p', 'An add-in for Excel for Windows (Microsoft 365, 2024, 2021, 2019, 2016). Microsoft says the Windows app has the full Power Query and Power Pivot features.', SRC.pqpp],
  ['VBA macros (.xlsm)', 'p', 'Excel for the web can open and edit a workbook with macros, but can\'t create, run or edit the macros; they stay in the file. Use Open in Desktop App.', SRC.vbaweb],
  ['PowerPoint Morph', 'p', 'Plays in Microsoft 365, 2019, 2021 and the web. PowerPoint 2016 plays it only with a recent build; otherwise it shows a Fade.', SRC.morph],
  ['PowerPoint Zoom', 'p', 'Created and played in Microsoft 365, 2019, 2021, 2024. In PowerPoint 2016 zooms play as plain hyperlinks with no zoom effect.', SRC.zoom]
];

// Changes to existing compatibility values, each confirmed on an official page (logged below)
const CHANGED = [];
const ver = n => d.versions.find(v => v.feature === n);
const py = ver('Python in Excel');
if (py && py.v2016 === 'n') { for (const k of ['v2016', 'v2019', 'v2021', 'v2024']) py[k] = 's'; CHANGED.push('Python in Excel in Office 2016–2024 now shows "View / static only": Microsoft says workbooks with Python can be viewed but Python cells show an error when recalculated (Microsoft Support: Python in Excel availability).'); }
const zm = ver('PowerPoint Zoom');
if (zm && zm.v2016 === 'n') { zm.v2016 = 'p'; zm.fails = 'Plays as plain hyperlinks (no zoom effect) in PowerPoint 2016; check web and Teams before presenting'; CHANGED.push('PowerPoint Zoom in PowerPoint 2016 corrected from "breaks / static picture" to "opens with limitations": Microsoft says zooms play as hyperlinks with no zooming transition (Microsoft Support: Use zoom for PowerPoint).'); }

// ---------- 8. File sharing guidance for Finance ----------
d.shareGuide = {
  formats: [
    ['.xlsx', 'Normal workbooks: schedules, working papers, data for others to use', 'No macros. Check Compatibility before sending to older Excel.'],
    ['.xlsm', 'Workbooks that contain VBA macros', 'Macros in files from email or the internet are blocked by default on most PCs; send .xlsx if the receiver doesn\'t need the macro.'],
    ['.csv', 'Plain data for uploads or other systems', 'One sheet, values only: no formulas, formatting or multiple sheets. Dates and leading zeros can change when opened.'],
    ['.pdf', 'Final reports and packs going to management, auditors or outside the company', 'Fixed layout that everyone can open and nobody edits by accident. Not for numbers people must reuse.'],
    ['.pptx', 'Presentations people may edit', 'Break links to Excel on the copy you send outside.'],
    ['.docx', 'Letters, memos and policies people may edit', 'Use Track Changes for review; send PDF for the final signed version.'],
    ['.xltx', 'Excel templates (e.g. a monthly forecast input file)', 'Opening it creates a new copy, so the template stays clean.'],
    ['.dotx', 'Word templates (letters, memos)', 'Same idea: every use starts a fresh document.']
  ],
  pdfWhy: 'Final financial reports are often shared as PDF when nobody needs to edit them: the layout is fixed, it opens on any device, and the numbers can\'t be changed or recalculated by accident. Keep the Excel version as your working file.',
  pq: [
    'Avoid sources on a personal path such as C:\\Users\\your-name\\Downloads\\… : the query breaks for everyone else and when your PC changes.',
    'Prefer a controlled team location such as a SharePoint or Teams library your team owns.',
    'Even on SharePoint, moving or renaming the folder or file can break the query, because it stores the path or URL. Agree a fixed folder and file name.',
    'Use a parameter for the folder path so it can be changed in one place (see the Parameters card).',
    'Write the source location in a short note (or SOP) next to the report.'
  ]
};

// ---------- 9. IT restrictions: governance level, what to check, copyable request ----------
const ITX = {
  'Macros from email / internet': ['MEDIUM', 'Is there a red or yellow message bar about macros under the ribbon? Did the file come from email or a website?', 'Please can you advise how I can run the macros in [file name], which we use for [purpose]? Can it be saved in a trusted SharePoint location, or can the macro be digitally signed?'],
  'Connected experiences (Designer, Translate, Dictate, Analyze Data, data types)': ['LOW', 'Is the feature greyed out or does it say it is turned off by your organisation?', 'Please can you confirm whether [feature] is allowed for my account? I would use it for [purpose] on internal, non-confidential content.'],
  'Office Add-ins store': ['MEDIUM', 'Does Get Add-ins show only "Admin Managed", or is the store missing?', 'Please can you review and, if approved, deploy the add-in [name] for my account? It is needed for [purpose].'],
  "'Anyone' sharing links, external sharing": ['HIGH', 'Which link types appear when you click Share? Is "Anyone" missing?', 'Please can you advise the approved way to share [folder] with our external auditors [firm], for example Specific people links or guest access, until [date]?'],
  'Auto-forward to external email': ['HIGH', 'Does your forwarding rule fail or bounce for outside addresses?', 'Please can you advise an approved way for [mailbox] to send [type of emails] to [recipient], for example a shared mailbox or an approved flow?'],
  'Publish to web, export data, Analyze in Excel': ['HIGH', 'Are Export or Analyze in Excel missing in the report menu?', 'Please can you confirm whether export or Analyze in Excel can be enabled for the report [name] for [people], for [purpose]?'],
  'On-premises data gateway': ['MEDIUM', 'Does scheduled refresh fail with a gateway message?', 'Please can you advise whether our report [name] can use the central data gateway for the source [server/file], with refresh [frequency]?'],
  'Premium connectors / DLP policies': ['MEDIUM', 'Does the flow show a premium (diamond) label or a DLP policy error when saving?', 'Please confirm whether the required Power Automate connector is permitted under our organisation\'s DLP policy and whether a Power Automate Premium licence is required.'],
  'Power Automate for desktop / external tools install': ['HIGH', 'Does the installer ask for admin rights?', 'Please can you install Power Automate for desktop on my PC? The intended use is [task]. It will run attended (while I am logged on) and will not post financial transactions.'],
  'SAP GUI scripting': ['HIGH', 'Does SAP show a scripting warning, or does a recorded script fail to start?', 'Please confirm whether SAP GUI scripting can be enabled for my user account. The intended use is read-only extraction of transaction data for a recurring Finance report. No financial posting activity will be automated.'],
  'Office Scripts, Python in Excel': ['MEDIUM', 'Is the Automate tab or the Python option missing in Excel?', 'Please can you confirm whether Office Scripts / Python in Excel can be enabled for my account? The intended use is [task] on internal data.'],
  'Copilot, AI uploads, web search in Copilot': ['MEDIUM', 'What label shows under your name (Copilot Chat (Basic), M365 Copilot (Basic) or (Premium))? Is file upload greyed out?', 'Please can you confirm which Copilot features I may use with internal Finance data, and whether a Microsoft 365 Copilot licence pilot is possible for [team]?'],
  'Meeting recording, transcription, audio': ['MEDIUM', 'Are Record and Transcribe greyed out in the meeting controls?', 'Please can you confirm the meeting policy for recording and transcription for [meeting type], and where recordings are stored?'],
  'Loop, external Teams access, apps in Teams': ['MEDIUM', 'Are Loop components or the app missing in Teams or Outlook?', 'Please can you advise whether [Loop / external access / app name] is available for our team, or planned?']
};
for (const b of d.itBlocks) { const x = ITX[b.what]; if (x && !b.level) { b.level = x[0]; b.check = x[1]; b.request = x[2]; } }
d.govLevels = [['LOW', 'Normal productivity feature'], ['MEDIUM', 'Shared or scheduled automation, or organisational data'], ['HIGH', 'External sharing, financial approvals, unattended automation, SAP automation, financial postings or sensitive data']];

// ---------- 10. Licence guidance: the six checks, and who needs which licence ----------
d.licLayers = [
  ['Microsoft supports it', 'The feature exists in your version and app (see the version on each card).'],
  ['A licence is needed', 'Some features need a paid plan or add-on (the Licence label on cards).'],
  ['Your company bought it', 'The company must have that licence in its agreement.'],
  ['It is assigned to you', 'IT assigns licences person by person. Check the label under your name for Copilot.'],
  ['IT has switched it on', 'Admins can turn features, connectors and agents off for everyone or for groups.'],
  ['Policy allows it', 'Security, DLP and AI policies can still limit what data you may use with it.']
];
d.licWho = [
  ['Microsoft 365 apps (Excel, Word, PowerPoint, Outlook, Teams)', 'Each user', 'Comes with your Microsoft 365 plan.', null],
  ['Power BI Free', 'Each user', 'Build in your own My workspace. Can view shared content only when both the report and its model are in Premium or Fabric F64+ capacity.', SRC.pbiUser],
  ['Power BI Pro', 'Creator and every viewer', 'Needed to publish to a workspace, edit and share, and to view content in Pro workspaces.', SRC.pbiUser],
  ['Power BI Premium Per User (PPU)', 'Everyone who opens PPU content', 'Content in a PPU workspace or app needs PPU to view; Pro users can\'t view it in the portal.', SRC.ppu],
  ['Microsoft Fabric capacity (F SKU)', 'Company (capacity)', 'On F64 or larger, Free users with a viewer role can view Power BI content. Below F64, every viewer needs Pro, PPU or a trial.', SRC.fabricLic],
  ['Power Automate Premium', 'Flow owner, or each user who runs it', 'Automated and scheduled flows run under the flow owner. Instant flows with premium connectors need Premium for every user who runs them (or a Process licence on the flow).', SRC.paFaq],
  ['Power Automate Process', 'One flow or one machine', 'For back-end automation not run by a person; on a machine it is an unattended bot running one desktop flow at a time.', SRC.paTypes],
  ['Power Apps Premium', 'Each app user', 'Needed for apps that use premium, custom or on-premises connectors, or Dataverse. Microsoft 365 rights cover standard connectors.', SRC.ppFaq],
  ['Microsoft 365 Copilot', 'Each Copilot user', 'Each user needs a Copilot licence plus a qualifying licence; no device-based licensing. Without it you get Copilot Chat.', SRC.copReq],
  ['Teams Premium', 'Organiser or each user, by feature', 'Some features need only the meeting organiser\'s licence (e.g. watermarking, meeting templates); intelligent recap features benefit only licensed users.', SRC.teamsPrem],
  ['Planner Premium (Planner Plan 1)', 'Plan creator', 'Needed to create premium plans. Microsoft 365 users can edit premium plans created by licensed users, without the advanced features.', SRC.plannerFaq]
];

// ---------- 11. Copilot: what can I use? ----------
d.copilotCheck = {
  note: 'Seeing a Copilot button does not mean every Copilot feature works for you. The button can open a limited chat on Basic, and IT can turn features off.',
  rows: [
    ['Can I ask general questions?', 'Yes (web answers, files you add)', 'Yes', 'Yes', 'FEAT-COPILOT-002'],
    ['Can Copilot use my work emails, chats, meetings and files?', 'No, only files you upload', 'Yes, with Work IQ on', 'Yes', 'FEAT-COPILOT-003'],
    ['Can it work inside Excel?', 'Limited or standard access, depending on your label', 'Yes, full', 'Yes, full', 'FEAT-COPILOT-019'],
    ['Can it work inside Word?', 'Limited, depending on your label', 'Yes, full', 'Yes, full', 'FEAT-COPILOT-018'],
    ['Can it summarise Outlook email?', 'Yes', 'Yes, and uses more of your work context', 'Yes', 'FEAT-COPILOT-006'],
    ['Can it summarise Teams meetings?', 'No', 'Yes (transcription must be on)', 'Yes', 'FEAT-COPILOT-013'],
    ['Can it work with Power BI?', 'Only if your reports sit on capacity with Copilot on', 'Same as Basic', 'Yes', 'FEAT-COPILOT-010'],
    ['Does it need another licence?', 'No paid Copilot licence', 'Microsoft 365 Copilot, per user', 'Plus Fabric F2+ or Premium P1+ capacity (company)', 'FEAT-COPILOT-001'],
    ['Could IT have turned it off?', 'Yes', 'Yes', 'Yes', 'FEAT-COPILOT-029']
  ],
  src: [SRC.copLic, SRC.xlCopFaq, SRC.copReq, SRC.pbiCop]
};
d.copilotCheck.rows.forEach(r => card(r[4]));

// ---------- 12. Finance data safety ----------
d.dataSafety = {
  never: 'NEVER type or paste passwords, PINs, authentication codes, security tokens, API keys or private keys into Copilot or any AI prompt, flow, macro or script.',
  note: 'Using Microsoft Copilot with a work account does not make every kind of data automatically safe to use. Your organisation\'s data classification, Microsoft 365, Copilot and security policies decide what is allowed.',
  rows: [
    ['Public information', 'Published annual reports, public price lists, published exchange rates', 'Generally fine to use; still check AI answers.', 'LOW'],
    ['Internal company information', 'Procedures, internal memos, team plans', 'Only in tools your company approves (for example Copilot signed in with your work account). Follow policy.', 'MEDIUM'],
    ['Customer data', 'Names, contacts, balances, credit limits', 'Only if policy allows; use the minimum columns needed.', 'HIGH'],
    ['Vendor data', 'Vendor master, terms, bank details', 'Only if policy allows; never bank account details in prompts.', 'HIGH'],
    ['Employee data', 'Staff lists, performance, personal details', 'Treat as confidential personal data; follow HR and privacy policy.', 'HIGH'],
    ['Payroll information', 'Salaries, deductions, payslips', 'Highly confidential: only approved systems and people.', 'HIGH'],
    ['Bank information', 'Account numbers, statements, balances', 'Highly confidential: do not paste account numbers into AI prompts.', 'HIGH'],
    ['Payment information', 'Payment files, approvals, card details', 'Highly confidential: payment files only in approved systems.', 'HIGH'],
    ['Credentials', 'Passwords, PINs, one-time codes', 'NEVER in any prompt, flow, macro, script or shared file.', 'NEVER'],
    ['API keys and security tokens', 'Keys, tokens, private keys, connection strings', 'NEVER in prompts or files; IT manages these.', 'NEVER'],
    ['Unpublished financial results', 'Draft results before announcement, forecasts', 'Restricted: only approved tools and named people.', 'HIGH'],
    ['SAP exports', 'GL, AR, AP, inventory extracts', 'Often contain customer, vendor or bank data: remove columns you don\'t need before sharing.', 'HIGH'],
    ['Audit information', 'PBC files, findings, auditor correspondence', 'Share only with the audit team through approved links.', 'HIGH'],
    ['Tax information', 'Returns, computations, correspondence', 'Confidential: approved tools and people only.', 'HIGH']
  ]
};

// ---------- 13. Self-learning cycle ----------
d.learnCycle = [
  ['Find', 'Find the feature: search, a learning path, or "Which tool for the job?".'],
  ['Try', 'Use it on sample data or a copy of a real file. Never the live file.'],
  ['Inspect', 'Look at what was produced: the formula, the Applied Steps, the model, the flow run.'],
  ['Explain', 'Say in your own words why it works. If you can\'t, open the Handbook page.'],
  ['Rebuild', 'Do it again from a blank sheet without the steps.'],
  ['Apply', 'Use it on a real Finance task this week.'],
  ['Document', 'Write the final method in a short SOP or reference note so others can run it.']
];

// ---------- 14. Licence ladder: Teams Premium events note (sourced) ----------
const tp = d.licences.find(l => l.name === 'Teams Premium');
if (tp && !tp.items.some(x => /April 2026/.test(x))) { tp.items.push('Note: Teams Premium bought after April 2026 no longer gives advanced webinar and town hall features; these follow the organiser\'s licence (Microsoft Learn: Teams events licensing overview)'); CHANGED.push('Teams Premium in the licence ladder: added Microsoft\'s note that Teams Premium bought after April 2026 no longer includes advanced webinar and town hall features (Microsoft Learn: Teams events licensing overview).'); }

// ---------- 15. Items flagged for manual review (kept as they are) ----------
d.review = [
  ['Sharing & old versions › XLOOKUP', 'Says "#NAME? after recalculation". Microsoft\'s XLOOKUP page confirms it is not in Excel 2016/2019 but does not describe what the other person sees. Kept; please re-check.', SRC.xlookup],
  ['Sharing & old versions › Data Model / Power Pivot (web)', 'Says "Web/Mac: view, not edit". The Microsoft pages read on 27 Nov 2026 did not confirm Mac or web behaviour. Kept; please re-check.', SRC.pqpp],
  ['Sharing & old versions › FILTER / SORT / UNIQUE in 2019', 'Marked "~". Microsoft says compatible formulas may show as legacy array formulas and newer functions may not calculate correctly. Kept; please re-check per function.', SRC.dyn],
  ['Card: Copilot inside Word / Excel / PowerPoint / OneNote', 'Overlaps the separate Copilot in Word, Excel, PowerPoint and OneNote cards. Kept both so existing links still work.', null],
  ['Update history › 3 Nov 2026', 'Says Copilot in Power BI needs F64+; corrected on 22 Nov 2026 to F2+ / P1+. The historical entry is kept unchanged.', null],
  ['Card: Data Model & Power Pivot (DAX in Excel)', 'Which Excel editions include Power Pivot was not confirmed on the Microsoft pages read. Please re-check the card\'s version text.', SRC.pivot]
];

// ---------- sources and update history ----------
for (const s of Object.values(SRC)) addSrc(d.sources, s);
if (!d.changes.some(c => c.date === T && c.items.some(x => /Finance upgrade/.test(x)))) {
  const items = [
    'Finance upgrade: new start choices on the home page (Solve a work problem, Learn an application, Follow a learning path), clearer progress wording, and sections grouped as Learn, Solve, Check and Reference.',
    'New Learning paths: five Finance levels (Excel foundation, Data preparation, Analysis & reporting, Automation, AI productivity) built from existing cards and Handbook pages, plus a "Recommended for Finance & Accounts" shortlist.',
    'Which tool for the job?: 18 new Finance scenarios (forecast consolidation, three-way variance, management pack, expense analysis, accrual reversals, AP ageing, inventory ageing, standard vs actual cost, realised and unrealised FX, intercompany, CAPEX, fixed assets, audit PBC, cash flow, payment documents, SAP-to-Excel, customer ageing), scale guidance and a tool comparison (Microsoft Learn licensing pages).',
    'Cards: "Try it" exercises with expected results on 26 core cards, and a related/next section; the how-to panel now follows one order.',
    'Search understands Finance words and abbreviations (AR, AP, P&L, BGT, FCT, FX, GL, COGS, month-end, SAP export, variance and more).',
    'Sharing & old versions: four compatibility levels and verified notes on dynamic arrays, XLOOKUP, LET, LAMBDA, VSTACK/HSTACK, Office Scripts, Python in Excel, Power Query, Power Pivot, macros, Morph and Zoom; file-format and Power Query source guidance (Microsoft Support and Microsoft Learn function and availability pages).',
    'IT blocks & licences: governance level, what you can check and copyable IT-request wording for each block; who needs which licence (Microsoft Learn); a Copilot "what can I use?" table; and Finance data safety guidance.',
    'Handbook: difficulty, estimated time, prerequisites, Finance use, practice exercises and next-lesson links on the main Excel, Power Query, Power BI and Power Automate pages; Interview tips now read "At work / Interview tip".',
    'Self-learn kit: the seven-step learning cycle (Find, Try, Inspect, Explain, Rebuild, Apply, Document).'
  ].concat(CHANGED);
  const today = d.changes.find(c => c.date === T);
  if (today) today.items.push(...items); else d.changes.unshift({ date: T, items });
}
d.changes = d.changes.slice(0, 24);
fs.writeFileSync('data.json', JSON.stringify(d, null, 1));
console.log('data.json migrated:', d.features.length, 'cards,', d.tasks.length, 'tasks,', CHANGED.length, 'value changes');
