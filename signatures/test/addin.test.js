// Runs the built add-in bundle in a bare V8 context (no window/document, like
// classic Outlook's JavaScript-only runtime) against a fake Office API.
const test = require('node:test');
const assert = require('node:assert');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const bundle = fs.readFileSync(path.join(__dirname, '..', 'addin', 'launchevent.js'), 'utf8');
const template = JSON.parse(fs.readFileSync(path.join(__dirname, '..', 'template.json'), 'utf8'));
const addressHtml = require('../src/renderer').escapeHtml(template.brand.address);

function loadAddin({ composeType = 'newMail', itemType = 'message', graph = null, templateOk = true } = {}) {
  const handlers = {};
  const calls = { signature: null, disabled: false, completed: 0, fetches: [] };
  const store = new Map();
  const ok = (value) => ({ status: 'succeeded', value });
  const item = {
    itemType,
    getComposeTypeAsync: (cb) => cb(ok({ composeType, coercionType: 'html' })),
    disableClientSignatureAsync: (cb) => { calls.disabled = true; cb(ok()); },
    body: { setSignatureAsync: (html, opts, cb) => { calls.signature = { html, opts }; cb(ok()); } }
  };
  const Office = {
    AsyncResultStatus: { Succeeded: 'succeeded', Failed: 'failed' },
    CoercionType: { Html: 'html' },
    MailboxEnums: { ItemType: { Message: 'message', Appointment: 'appointment' } },
    onReady: () => Promise.resolve(),
    actions: { associate: (name, fn) => { handlers[name] = fn; } },
    context: { mailbox: { item, userProfile: { displayName: 'Alex Morgan', emailAddress: 'Alex.Morgan@waterbabies.co.uk' } } }
  };
  const fetch = async (url) => {
    calls.fetches.push(url);
    if (url.includes('/template.json')) {
      return templateOk ? { ok: true, json: async () => template } : { ok: false, status: 500 };
    }
    if (graph && url.startsWith('https://graph.microsoft.com/')) return { ok: true, json: async () => graph };
    throw new Error('unexpected fetch ' + url);
  };
  const OfficeRuntime = { storage: { getItem: async (k) => store.get(k) ?? null, setItem: async (k, v) => { store.set(k, v); } } };
  const context = vm.createContext({ Office, OfficeRuntime, fetch, console: process.env.DEBUG ? console : { log() {}, warn() {}, error() {} }, setTimeout, clearTimeout, Promise, URL, TextEncoder, TextDecoder, crypto: globalThis.crypto });
  vm.runInContext(bundle, context);
  const run = (name) => new Promise((resolve) => handlers[name]({ completed: () => { calls.completed++; resolve(); } }));
  return { handlers, calls, run, store };
}

test('registers all handlers without a DOM', () => {
  const { handlers } = loadAddin();
  assert.deepStrictEqual(Object.keys(handlers).sort(), ['insertSignature', 'onNewAppointmentComposeHandler', 'onNewMessageComposeHandler']);
});

test('new email gets the full signature and Outlook signature is disabled', async () => {
  const { run, calls } = loadAddin();
  await run('onNewMessageComposeHandler');
  assert.strictEqual(calls.completed, 1);
  assert.ok(calls.disabled);
  assert.strictEqual(calls.signature.opts.coercionType, 'html');
  // Graph is unavailable here, so it falls back to the mailbox profile.
  assert.ok(calls.signature.html.includes('Alex Morgan'));
  assert.ok(calls.signature.html.includes(addressHtml));
});

test('reply gets the short signature', async () => {
  const { run, calls } = loadAddin({ composeType: 'reply' });
  await run('onNewMessageComposeHandler');
  assert.ok(!calls.signature.html.includes(addressHtml));
});

test('event always completes even if the template cannot load', async () => {
  const { run, calls } = loadAddin({ templateOk: false });
  await run('onNewMessageComposeHandler');
  assert.strictEqual(calls.completed, 1);
  assert.strictEqual(calls.signature, null);
});

test('template is cached between composes', async () => {
  const { run, calls } = loadAddin();
  await run('onNewMessageComposeHandler');
  await run('onNewMessageComposeHandler');
  assert.strictEqual(calls.fetches.filter((u) => u.includes('template.json')).length, 1);
  await run('insertSignature');
  assert.strictEqual(calls.fetches.filter((u) => u.includes('template.json')).length, 2, 'button should force a refresh');
});
