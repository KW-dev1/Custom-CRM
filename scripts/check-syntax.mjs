// Static check: every JS file must parse, and the manifest must be valid JSON.
import { readdirSync, statSync, readFileSync } from 'node:fs';
import { execFileSync } from 'node:child_process';
import { join } from 'node:path';

const walk = (d) => readdirSync(d).flatMap((f) => {
  if (f === 'node_modules' || f.startsWith('.')) return [];
  const p = join(d, f);
  return statSync(p).isDirectory() ? walk(p) : [p];
});
let failed = false;
for (const f of walk('extension').filter((f) => /\.(m?js)$/.test(f))) {
  try { execFileSync(process.execPath, ['--check', f], { stdio: 'pipe' }); }
  catch (e) { failed = true; console.error(`SYNTAX ERROR: ${f}\n${e.stderr}`); }
}
JSON.parse(readFileSync('extension/manifest.json', 'utf8'));
if (failed) process.exit(1);
console.log('syntax check ok');
