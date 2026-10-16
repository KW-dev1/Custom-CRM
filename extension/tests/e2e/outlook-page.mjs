// Simulated Outlook Web page used by the real-Chromium e2e scripts (NOT the real Outlook).
// Mirrors the structure captured from a real Outlook Web compose window: To/Cc are contenteditable divs
// (no role=textbox), resolved recipients are span[contenteditable=false][aria-label=address] INSIDE the field,
// Cc/Bcc are plain buttons, and Enter or ";" turns the typed text into a chip.
export const PAGE = `<!doctype html><html><body><div role="main"><button aria-label="New mail">New mail</button></div>
<div role="tree"><div role="treeitem" data-folder-name="inbox" aria-selected="true"><span>Inbox</span></div><div role="treeitem" data-folder-name="sent items" aria-selected="false"><span>Sent Items</span></div></div>
<div role="region"><div role="listbox" aria-label="Message list" id="msglist"></div></div>
<script>
window.__sent = [];
window.__folderClicks = [];
const mb = { cur: 'inbox', sent: [] };
const pad = (n) => String(n).padStart(2, '0');
const clock = (d) => ((d.getHours() + 11) % 12 + 1) + ':' + pad(d.getMinutes()) + ' ' + (d.getHours() < 12 ? 'AM' : 'PM');
const full = (d) => ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'][d.getDay()] + ' ' + (d.getMonth() + 1) + '/' + d.getDate() + '/' + d.getFullYear() + ' ' + clock(d);
function renderList() {
  const list = document.getElementById('msglist'); list.innerHTML = '';
  const rows = mb.cur === 'inbox' ? [{ who: 'news@example.com', subject: 'Weekly digest', at: new Date(Date.now() - 3600e3) }] : mb.sent;
  rows.forEach((m) => { const r = document.createElement('div'); r.setAttribute('role', 'option');
    r.setAttribute('aria-label', m.who + ' ' + m.subject + ' ' + clock(m.at) + ' preview');
    r.innerHTML = '<div><span></span><span></span><span class="t"></span><span>preview</span></div>';
    const sp = r.querySelectorAll('span'); sp[0].textContent = m.who; sp[1].textContent = m.subject; sp[2].textContent = clock(m.at); sp[2].setAttribute('title', full(m.at));
    list.appendChild(r); });
}
document.querySelectorAll('[role=treeitem]').forEach((it) => it.addEventListener('click', () => {
  mb.cur = it.getAttribute('data-folder-name'); window.__folderClicks.push(mb.cur);
  document.querySelectorAll('[role=treeitem]').forEach((x) => x.setAttribute('aria-selected', String(x === it))); renderList(); }));
renderList();
document.querySelector('[aria-label="New mail"]').addEventListener('click', () => setTimeout(() => {
  const w = document.createElement('div'); w.id = 'compose';
  w.innerHTML = '<div><button aria-label="Send" type="button">Send</button><button aria-label="Discard" type="button">Discard</button></div>'
    + '<div><button type="button" role="button">To</button><div role="group"><div aria-label="To" contenteditable="true" tabindex="0"></div></div><button type="button">Cc</button><button type="button">Bcc</button></div>'
    + '<div role="group"><div aria-label="Cc" contenteditable="true" tabindex="0" hidden></div></div>'
    + '<input aria-label="Subject" placeholder="Add a subject" type="text"><div role="textbox" aria-label="Message body" contenteditable="true" tabindex="0"></div><div role="alert"></div>';
  document.body.appendChild(w);
  const ccToggle = [...w.querySelectorAll('button')].find((b) => b.textContent === 'Cc');
  ccToggle.addEventListener('click', () => { w.querySelector('[aria-label="Cc"][contenteditable]').hidden = false; });
  const typed = (f) => [...f.childNodes].filter((n) => n.nodeType === 3).map((n) => n.data).join('');
  const resolve = (f) => { const t = typed(f).replace(/;$/, '').trim(); if (!t) return;
    [...f.childNodes].filter((n) => n.nodeType === 3).forEach((n) => n.remove());
    const chip = document.createElement('span'); chip.setAttribute('contenteditable', 'false'); chip.setAttribute('aria-label', t); chip.textContent = t; f.appendChild(chip); };
  for (const label of ['To', 'Cc']) {
    const f = w.querySelector('[aria-label="' + label + '"][contenteditable]');
    f.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); resolve(f); } });
    f.addEventListener('input', () => { if (typed(f).endsWith(';')) resolve(f); });
  }
  w.querySelector('button[aria-label="Send"]').addEventListener('click', () => {
    const chips = (label) => [...w.querySelectorAll('[aria-label="' + label + '"][contenteditable] span[contenteditable="false"]')].map((x) => x.getAttribute('aria-label'));
    setTimeout(() => { mb.sent.unshift({ who: chips('To')[0], subject: w.querySelector('input').value, at: new Date() }); if (mb.cur === 'sent items') renderList(); }, 300);
    window.__sent.push({ to: chips('To'), cc: chips('Cc'), subject: w.querySelector('input').value, body: w.querySelector('[aria-label^="Message body"]').innerText });
    setTimeout(() => w.remove(), 50);
  });
}, 20));
</script></body></html>`;
