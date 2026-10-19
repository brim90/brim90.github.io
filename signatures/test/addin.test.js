// Runs the built add-in in a bare V8 context (no window or document, like
// classic Outlook's JavaScript-only runtime) against a fake Outlook.
const test = require('node:test');
const assert = require('node:assert');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const root = path.join(__dirname, '..');
const bundle = fs.readFileSync(path.join(root, 'addin', 'launchevent.js'), 'utf8');

function loadAddin({ serverUp = true, mobile = false } = {}) {
  const handlers = {};
  const calls = { signature: null, disabled: false, completed: 0, fetches: [] };
  const store = new Map();
  const ok = (value) => ({ status: 'succeeded', value });
  const item = {
    body: { setSignatureAsync: (html, opts, cb) => { calls.signature = { html, opts }; cb(ok()); } }
  };
  // Outlook mobile may not offer every 1.10 API; the add-in must cope.
  if (!mobile) item.disableClientSignatureAsync = (cb) => { calls.disabled = true; cb(ok()); };
  const Office = {
    AsyncResultStatus: { Succeeded: 'succeeded', Failed: 'failed' },
    CoercionType: { Html: 'html' },
    onReady: () => Promise.resolve(),
    actions: { associate: (name, fn) => { handlers[name] = fn; } },
    context: { mailbox: { item, userProfile: { displayName: 'Dan Brimmicombe', emailAddress: 'Dan@waterbabies.co.uk' } } }
  };
  const fetch = async (url) => {
    calls.fetches.push(url);
    const file = url.replace(/^https:\/\/danbrimmicombe\.co\.uk\/signatures\//, '').replace(/\?.*$/, '');
    if (!serverUp || file === url) return { ok: false, status: 503 };
    const text = fs.readFileSync(path.join(root, file), 'utf8');
    return { ok: true, text: async () => text };
  };
  const OfficeRuntime = { storage: { getItem: async (k) => store.get(k) ?? null, setItem: async (k, v) => { store.set(k, v); } } };
  const quiet = { log() {}, warn() {}, error() {} };
  const context = vm.createContext({ Office, OfficeRuntime, fetch, console: process.env.DEBUG ? console : quiet, setTimeout, clearTimeout, URL, TextEncoder, TextDecoder, crypto: globalThis.crypto });
  vm.runInContext(bundle, context);
  const run = () => new Promise((resolve) => handlers.onNewMessageComposeHandler({ completed: () => { calls.completed++; resolve(); } }));
  return { handlers, calls, run };
}

test('registers the compose handler without a DOM', () => {
  assert.deepStrictEqual(Object.keys(loadAddin().handlers), ['onNewMessageComposeHandler']);
});

test('inserts the signature and turns off Outlook\'s own one', async () => {
  const { run, calls } = loadAddin();
  await run();
  assert.strictEqual(calls.completed, 1);
  assert.ok(calls.disabled);
  assert.strictEqual(calls.signature.opts.coercionType, 'html');
  // No Graph in the test, so it falls back to the mailbox name.
  assert.ok(calls.signature.html.includes('Dan Brimmicombe'));
  assert.ok(calls.signature.html.includes('https://danbrimmicombe.co.uk/signatures/assets/logo.png'));
});

test('works when disableClientSignatureAsync is unavailable (mobile)', async () => {
  const { run, calls } = loadAddin({ mobile: true });
  await run();
  assert.ok(calls.signature.html.includes('Dan Brimmicombe'));
});

test('always completes the event, even when the site is down', async () => {
  const { run, calls } = loadAddin({ serverUp: false });
  await run();
  assert.strictEqual(calls.completed, 1);
  assert.strictEqual(calls.signature, null);
});

test('designs are cached between emails', async () => {
  const { run, calls } = loadAddin();
  await run();
  await run();
  assert.strictEqual(calls.fetches.filter((u) => u.includes('signatures.json')).length, 1);
});
