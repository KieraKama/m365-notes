// Compares the current JSON files with a saved copy and reports anything that was removed or changed.
// Additions (new keys, new array items at the end, new items in lists) are allowed.
// Usage: node tools/check-integrity.js <folder-with-original-json>
const fs = require('fs');
const path = require('path');
const dir = process.argv[2];
if (!dir) { console.error('Usage: node tools/check-integrity.js <folder-with-original-json>'); process.exit(2); }
const files = ['data.json', 'handbook.json', 'ribbon.json', 'formulas.json', 'filetypes.json'];
// Paths whose change is intended and logged in data.json "changes".
const ALLOWED = [
  /^data\.json:\.?versions\[(8|17)\]\.(v2016|v2019|v2021|v2024|fails)$/, // Python in Excel, PowerPoint Zoom (logged 27 Nov 2026)
  /^data\.json:\.?features\[\d+\]\.updated$/,                          // card got new content (Try it) on 27 Nov 2026
  /^data\.json:\.?changes/
];
const problems = [];
function walk(a, b, p) {
  if (a === null || typeof a !== 'object') {
    if (a !== b && !ALLOWED.some(r => r.test(p))) problems.push(`${p}: changed ${JSON.stringify(a).slice(0, 80)} -> ${JSON.stringify(b).slice(0, 80)}`);
    return;
  }
  if (b === null || typeof b !== 'object') { problems.push(`${p}: removed or replaced`); return; }
  if (Array.isArray(a)) {
    if (!Array.isArray(b)) { problems.push(`${p}: no longer a list`); return; }
    if (p.endsWith('changes')) { // history: every old entry must still be present, unchanged
      const old = a.map(x => JSON.stringify(x)), now = b.map(x => JSON.stringify(x));
      const todayOld = a.filter(x => !now.includes(JSON.stringify(x)));
      todayOld.forEach(x => { const m = b.find(y => y.date === x.date); if (!m || !x.items.every(i => m.items.includes(i))) problems.push(`${p}: history entry ${x.date} lost items`); });
      return;
    }
    a.forEach((x, i) => walk(x, b[i], `${p}[${i}]`));
    return;
  }
  for (const k of Object.keys(a)) {
    if (!(k in b)) { problems.push(`${p}.${k}: key removed`); continue; }
    walk(a[k], b[k], p ? `${p}.${k}` : k);
  }
}
for (const f of files) {
  const a = JSON.parse(fs.readFileSync(path.join(dir, f), 'utf8'));
  const b = JSON.parse(fs.readFileSync(f, 'utf8'));
  walk(a, b, f + ':');
}
const fixP = s => s.replace(':.', ':');
if (problems.length) { console.log('PROBLEMS (' + problems.length + '):\n' + problems.map(fixP).join('\n')); process.exit(1); }
console.log('Integrity OK: nothing removed or changed outside the allowed, logged changes.');
