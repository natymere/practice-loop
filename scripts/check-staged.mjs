import {execFileSync} from 'node:child_process';

// Read the index, not working files, and never print a matched value.
const git = (...args) => execFileSync('git', args, {encoding:'utf8', maxBuffer:8 * 1024 * 1024});
const files = git('diff','--cached','--name-only','--diff-filter=ACMR','-z').split('\0').filter(Boolean);
const secretPatterns = [
  /-----BEGIN (?:RSA |EC |OPENSSH |ENCRYPTED )?PRIVATE KEY-----/,
  /\b\d{6,12}:[A-Za-z0-9_-]{30,}\b/,
  /\bAIza[A-Za-z0-9_-]{35}\b/,
  /["'](?:refresh_token|client_secret|private_key)["']\s*:\s*["'][^"']+["']/,
];
let failed = false;
for (const file of files) {
  const localFile = /(^|\/)(?:\.clasprc[^/]*|\.env(?:\.[^/]*)?|secrets|private|local|backups)(?:\/|$)/.test(file)
    || /(^|\/)(?:docs|\.agents|\.codex|\.claude|\.cursor)(?:\/|$)/.test(file)
    || /\.(?:md|markdown)$/i.test(file) && file !== 'README.md'
    || /(^|\/)\.clasp(?:\.[^/]+)?\.json$/.test(file) && file !== '.clasp.example.json'
    || /(?:credentials|service-account|token|properties)[^/]*\.json$/i.test(file)
    || /\.(?:pem|key|p12|pfx|xlsx|xls)$/i.test(file);
  if (localFile) { console.error(`Blocked local/private file: ${file}`); failed = true; continue; }
  const contents = git('show',`:${file}`);
  const literalProperty = file.startsWith('src/') && /TELEGRAM_(?:BOT_TOKEN|CHAT_ID)["']?\s*[:=]\s*["'][^"']+["']/.test(contents);
  if (literalProperty || secretPatterns.some(pattern => pattern.test(contents))) {
    console.error(`Possible credential in ${file}. Keep values outside Git.`); failed = true;
  }
}
try {
  const manifest = JSON.parse(git('show',':src/appsscript.json'));
  if (manifest.webapp?.access !== 'MYSELF' || manifest.webapp?.executeAs !== 'USER_DEPLOYING') throw new Error();
} catch {
  console.error('The staged manifest must keep owner-only deployment.'); failed = true;
}
if (failed) process.exit(1);
console.log('Staged-file credential and private-deployment checks passed.');
