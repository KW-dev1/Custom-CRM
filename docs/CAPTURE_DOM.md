# Capturing a mail page's structure

The extension drives a webmail page through its HTML. To add a new site (Yahoo Mail) or to fix a selector after a site changes,
the maintainer needs real samples of that page's structure. These two console snippets collect **only element names and
attributes** (tag, role, aria-label, data-test-id, ...). Email addresses are replaced with `<email>` and long ids with `<id>`.
They read the page only; they change nothing and send nothing. Skim the output before sharing it.

How to run a snippet: open the page, press **F12**, choose the **Console** tab, paste the code, press Enter. If Chrome refuses to
paste, type `allow pasting` and press Enter first. The result is copied to your clipboard; paste it into the chat.

## Snippet A: the compose window

Take it at each of these moments (label them A1, A2, ...):

1. **A1** an empty new message (click Compose/New message, type nothing)
2. **A2** after you typed your own address into **To** and pressed Enter (or `;`), so it became a chip
3. **A3** after you clicked the **Cc** control so the Cc field is showing (nothing typed in it)
4. **A4** after you typed a subject and a two-line message body (shows how the editor stores paragraphs)

```js
(() => {
  const vis = (el) => { const r = el.getBoundingClientRect(); return r.width > 0 && r.height > 0; };
  const attr = (el, a) => el.getAttribute(a) || '';
  const subj = [...document.querySelectorAll('input,textarea')].find((el) => vis(el) && /subject/i.test(['aria-label', 'placeholder', 'name', 'data-test-id', 'id'].map((a) => attr(el, a)).join(' ')));
  if (!subj) { console.log('NO SUBJECT FIELD FOUND: open a NEW MESSAGE first'); return; }
  const editors = [...document.querySelectorAll('[contenteditable="true"]')].filter(vis).sort((a, b) => b.offsetWidth * b.offsetHeight - a.offsetWidth * a.offsetHeight);
  const body = editors[0];
  let root = subj;
  if (body) { while (root && !root.contains(body)) root = root.parentElement; }
  for (let i = 0; i < 3 && root && root.parentElement; i++) root = root.parentElement;
  const keep = ['role', 'aria-label', 'title', 'placeholder', 'contenteditable', 'type', 'name', 'id', 'data-test-id', 'data-testid', 'tabindex', 'aria-haspopup', 'aria-expanded', 'aria-invalid', 'aria-hidden', 'aria-selected', 'aria-checked', 'href'];
  const red = (s) => String(s).replace(/[^\s;,<>"]+@[^\s;,<>"]+/g, '<email>').replace(/[A-Za-z0-9_%-]{20,}/g, '<id>').slice(0, 80);
  const describe = (el) => { const o = { tag: el.tagName.toLowerCase() }; keep.forEach((a) => { if (el.hasAttribute(a)) o[a] = red(el.getAttribute(a)); }); const t = el.children.length === 0 ? (el.textContent || '').trim() : ''; if (t) o.text = red(t); return o; };
  const parentOf = (el) => { const p = el.parentElement; return p ? p.tagName.toLowerCase() + (p.getAttribute('role') ? `[role=${p.getAttribute('role')}]` : '') + (p.getAttribute('aria-label') ? `[aria-label=${red(p.getAttribute('aria-label'))}]` : '') + (p.getAttribute('data-test-id') ? `[data-test-id=${red(p.getAttribute('data-test-id'))}]` : '') : ''; };
  const rows = [];
  root.querySelectorAll('[contenteditable],input,textarea,button,a,select,[role],[aria-label],[data-test-id],[title]').forEach((el) => {
    if (!vis(el)) return;
    rows.push({ ...describe(el), parent: parentOf(el) });
  });
  const out = rows.slice(0, 250);
  copy(JSON.stringify(out, null, 1));
  console.log(`${out.length} elements copied to the clipboard. Skim before pasting.`);
})();
```

## Snippet B: a message in the Sent folder

1. Send yourself (or any address you own) a test email with an easy-to-spot subject, for example `zz capture test 1`.
2. Open the **Sent** folder so that email is listed.
3. In the snippet, replace `PUT YOUR TEST EMAIL SUBJECT HERE` with that subject.
4. Run it.

```js
(() => {
  const SUBJECT = 'PUT YOUR TEST EMAIL SUBJECT HERE';
  const vis = (el) => { const r = el.getBoundingClientRect(); return r.width > 0 && r.height > 0; };
  const keep = ['role', 'aria-label', 'title', 'aria-selected', 'aria-level', 'aria-posinset', 'data-test-id', 'data-testid', 'data-convid', 'data-folder-name', 'href', 'tabindex', 'type', 'name'];
  const red = (s) => String(s).replace(/[^\s;,<>"]+@[^\s;,<>"]+/g, '<email>').replace(/[A-Za-z0-9_%-]{24,}/g, '<id>').slice(0, 160);
  const describe = (el) => { const o = { tag: el.tagName.toLowerCase() }; keep.forEach((a) => { if (el.hasAttribute(a)) o[a] = red(el.getAttribute(a)); }); const t = el.children.length === 0 ? (el.textContent || '').trim() : ''; if (t) o.text = red(t); return o; };
  const chain = (el, n) => { const a = []; let p = el.parentElement; for (let i = 0; i < n && p; i++, p = p.parentElement) a.push(describe(p)); return a; };
  const out = { url: red(location.pathname), folder: [], subjectFound: false, subjectElement: null, ancestors: [], rowGuess: null, rowText: '', rowDescendants: [] };

  // 1) the Sent folder control(s)
  [...document.querySelectorAll('a,button,li,div,span,[role]')]
    .filter((el) => vis(el) && el.children.length < 4 && /^sent( items| mail)?$/i.test((el.getAttribute('aria-label') || el.getAttribute('title') || el.textContent || '').trim()))
    .slice(0, 6).forEach((el) => out.folder.push({ el: describe(el), parents: chain(el, 4) }));

  // 2) the message row that contains your test subject
  const leaf = [...document.querySelectorAll('body *')].find((el) => vis(el) && el.children.length === 0 && SUBJECT.length > 3 && (el.textContent || '').includes(SUBJECT));
  if (leaf) {
    out.subjectFound = true;
    out.subjectElement = describe(leaf);
    out.ancestors = chain(leaf, 9);
    let row = leaf;
    for (let i = 0; i < 9 && row.parentElement; i++) {
      const hint = [row.getAttribute('role'), row.getAttribute('data-test-id'), row.tagName].join(' ');
      if (/option|row|listitem|message|\bLI\b|\bA\b/i.test(hint) && row !== leaf) break;
      row = row.parentElement;
    }
    out.rowGuess = describe(row);
    out.rowText = [...row.querySelectorAll('*')].filter((el) => el.children.length === 0 && (el.textContent || '').trim()).map((el) => red(el.textContent.trim())).join(' | ');
    out.rowDescendants = [...row.querySelectorAll('*')].filter(vis).slice(0, 50).map(describe);
  }
  copy(JSON.stringify(out, null, 1));
  console.log(`Copied. Sent folder matches: ${out.folder.length}; subject found: ${out.subjectFound}. Skim before pasting.`);
})();
```

If the subject is **not found**, check you replaced the placeholder with the exact subject and that the Sent folder is open and showing it.

## Things worth telling the maintainer as well

- What happens when you type an address in **To** and press Enter, or `;`, or Tab: does it become a chip?
- Does clicking **Cc** reveal a new field, or open something else?
- Does the compose window close after Send? Is there a "message sent" notice?
- Are you using the new Yahoo Mail layout? (English UI, `mail.yahoo.com`.)
