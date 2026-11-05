// Rebuilds the folded "Site index" at the bottom of index.html from the JSON files.
// Run: node build-index.js [path/to/index.html]
const fs = require('fs');
const file = process.argv[2] || 'index.html';
const esc = s => String(s).replace(/[&<>"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const d = JSON.parse(fs.readFileSync('data.json','utf8'));
const hb = JSON.parse(fs.readFileSync('handbook.json','utf8'));
const fx = JSON.parse(fs.readFileSync('formulas.json','utf8'));
const rb = JSON.parse(fs.readFileSync('ribbon.json','utf8'));
const ft = JSON.parse(fs.readFileSync('filetypes.json','utf8'));
const apps = Object.fromEntries(d.apps.map(a => [a.id, a.name]));
let h = '<details class="site-index" id="site-index"><summary>Site index: everything on this site</summary><div class="si-grid">';
h += '<section><h3>Feature notes</h3>' + d.apps.map(a => {
  const list = d.features.filter(f => f.app === a.id);
  return list.length ? `<p><b>${esc(a.name)}:</b> ${list.map(f => `<a href="#notes" data-siq="${esc(f.name)}">${esc(f.name)}</a>`).join(' · ')}</p>` : '';
}).join('') + '</section>';
h += '<section><h3>Common questions</h3><p>' + (d.faq||[]).map(f => `<a href="#faq" data-sifaq="${esc(f.q)}">${esc(f.q)}</a>`).join(' · ') + '</p></section>';
h += '<section><h3>Handbook</h3>' + hb.subjects.map((s, si) => `<p><b>${esc(s.name)}:</b> ${s.pages.map((p, pi) => `<a href="#handbook" data-sihb="${si}|${pi}">${esc(p.t)}</a>`).join(' · ')}</p>`).join('') + '</section>';
h += '<section><h3>Excel formulas</h3><p><b>New functions:</b> ' + fx.fns.map(f => `<a href="#formulas" data-sifx="${esc(f[0])}">${esc(f[0])}</a>`).join(' · ') + '</p>';
h += '<p><b>Old → new:</b> ' + fx.swaps.map(s => `<a href="#formulas" data-sifx="${esc(s[1].split(' ')[0])}">${esc(s[0])} → ${esc(s[1])}</a>`).join(' · ') + '</p></section>';
h += '<section><h3>Start here: basics (button guide)</h3><p>' + rb.apps.map(a => `<a href="#basics">${esc(a.name)}</a> (${a.tabs.map(t => esc(t.name)).join(', ')})`).join(' · ') + '</p></section>';
h += '<section><h3>Which tool for the job?</h3><p>' + d.tasks.map(t => `<a href="#tools">${esc(t.task)}</a>`).join(' · ') + '</p></section>';
h += '<section><h3>Save as types</h3>' + Object.entries(ft.apps).map(([k, rows]) => `<p><b>${esc(apps[k] || k)}:</b> ${rows.map(r => esc(r[0].replace(' ★','') + ' ' + r[1])).join(' · ')}</p>`).join('') + '</section>';
h += '</div></details>';
let s = fs.readFileSync(file, 'utf8');
const A = '<!--SITE-INDEX:START-->', B = '<!--SITE-INDEX:END-->';
if (!s.includes(A)) throw new Error('markers missing in ' + file);
s = s.slice(0, s.indexOf(A) + A.length) + '\n' + h + '\n' + s.slice(s.indexOf(B));
fs.writeFileSync(file, s);
console.log('site index rebuilt in', file, '-', h.length, 'chars');
