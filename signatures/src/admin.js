/*
 * Admin console: signs an administrator in, reads every staff member from
 * Microsoft Graph, previews their signature against the live template, and
 * exports the template (for the Outlook add-in) or rendered signatures (for
 * the PowerShell server-side push).
 */
const { PublicClientApplication } = require('@azure/msal-browser');
const SignatureRenderer = require('./renderer');
const config = require('./config');

const SCOPES = ['User.Read.All'];
const GRAPH_USERS = 'https://graph.microsoft.com/v1.0/users?$top=999&$filter=accountEnabled eq true&$select=' + SignatureRenderer.GRAPH_SELECT;

const SAMPLE_USERS = [
  {
    id: 'sample-1', displayName: 'Alex Morgan', givenName: 'Alex', surname: 'Morgan',
    jobTitle: 'Head of Swim Programmes', department: 'Operations', companyName: 'Water Babies',
    businessPhones: ['0330 223 6777'], mobilePhone: '07700 900123',
    mail: 'alex.morgan@waterbabies.co.uk', onPremisesExtensionAttributes: {}
  },
  {
    id: 'sample-2', displayName: 'Sam Patel', givenName: 'Sam', surname: 'Patel',
    jobTitle: 'Franchise Support Manager, UK & EU', department: 'Franchise Support',
    businessPhones: [], mobilePhone: '07700 900456', mail: 'sam.patel@waterbabies.co.uk',
    onPremisesExtensionAttributes: {}
  },
  {
    id: 'sample-3', displayName: 'Jordan Lee', givenName: 'Jordan', surname: 'Lee',
    jobTitle: '', department: 'Finance', businessPhones: [],
    mail: 'jordan.lee@waterbabies.co.uk', onPremisesExtensionAttributes: {}
  }
];

const state = {
  template: null,
  users: SAMPLE_USERS,
  selectedId: SAMPLE_USERS[0].id,
  variant: 'new',
  previewDate: new Date().toISOString().slice(0, 10),
  filter: '',
  demo: true,
  account: null
};

const $ = (id) => document.getElementById(id);
let pca;

function setStatus(message, kind) {
  const el = $('status');
  el.textContent = message;
  el.className = 'status' + (kind ? ' ' + kind : '');
}

function previewOptions() {
  return { variant: state.variant, now: new Date(state.previewDate + 'T12:00:00') };
}

function selectedUser() {
  return state.users.find((u) => u.id === state.selectedId) || state.users[0];
}

function renderUserList() {
  const list = $('user-list');
  const q = state.filter.toLowerCase();
  const users = state.users.filter((u) =>
    [u.displayName, u.mail, u.department, u.jobTitle].some((v) => (v || '').toLowerCase().includes(q)));
  list.replaceChildren();
  for (const user of users) {
    const li = document.createElement('li');
    const button = document.createElement('button');
    button.type = 'button';
    button.className = 'user' + (user.id === state.selectedId ? ' selected' : '');
    const name = document.createElement('span');
    name.className = 'user-name';
    name.textContent = user.displayName || user.mail;
    const meta = document.createElement('span');
    meta.className = 'user-meta';
    meta.textContent = [user.jobTitle, user.department].filter(Boolean).join(' · ') || user.mail;
    button.append(name, meta);
    if (state.template) {
      const missing = SignatureRenderer.missingFields(state.template, user);
      if (missing.length) {
        const badge = document.createElement('span');
        badge.className = 'badge';
        badge.textContent = missing.length + ' missing';
        badge.title = 'Empty in Entra ID: ' + missing.join(', ');
        button.append(badge);
      }
    }
    button.addEventListener('click', () => {
      state.selectedId = user.id;
      renderUserList();
      renderPreview();
    });
    li.append(button);
    list.append(li);
  }
  $('user-count').textContent = users.length + ' of ' + state.users.length + (state.demo ? ' sample users' : ' users');
}

function renderPreview() {
  const user = selectedUser();
  if (!state.template || !user) return;
  let html;
  try {
    html = SignatureRenderer.renderSignature(state.template, user, previewOptions());
    $('template-error').textContent = '';
  } catch (e) {
    $('template-error').textContent = e.message;
    return;
  }
  const doc = '<!doctype html><html><head><meta charset="utf-8"></head><body style="margin:16px;background:#fff;">' +
    '<p style="font-family:Arial,sans-serif;font-size:13px;color:#333;margin:0 0 16px;">Hi,<br><br>Thanks for your email.<br><br>Kind regards,</p>' +
    html + '</body></html>';
  $('preview').srcdoc = doc;
  $('html-output').value = html;
  const missing = SignatureRenderer.missingFields(state.template, user);
  $('missing').textContent = missing.length
    ? 'Not set in Entra ID for this person (hidden in the signature): ' + missing.join(', ')
    : 'All fields used by the template are filled in for this person.';
  $('missing').className = 'note' + (missing.length ? ' warn' : '');
}

function brandFields() {
  const brand = state.template.brand || {};
  return Object.keys(brand).filter((k) => typeof brand[k] === 'string' || typeof brand[k] === 'number');
}

function renderBrandForm() {
  const form = $('brand-form');
  form.replaceChildren();
  for (const key of brandFields()) {
    const label = document.createElement('label');
    const caption = document.createElement('span');
    caption.textContent = key.replace(/([A-Z])/g, ' $1').replace(/^./, (c) => c.toUpperCase());
    const value = state.template.brand[key];
    const isColor = /color$/i.test(key) && /^#[0-9a-f]{6}$/i.test(value);
    const input = document.createElement(key === 'legal' ? 'textarea' : 'input');
    if (isColor) input.type = 'color';
    else if (typeof value === 'number') input.type = 'number';
    input.value = value;
    input.addEventListener('input', () => {
      state.template.brand[key] = typeof value === 'number' ? Number(input.value) : input.value;
      syncJson();
      refresh();
    });
    label.append(caption, input);
    form.append(label);
  }
}

function syncJson() {
  $('template-json').value = JSON.stringify(state.template, null, 2);
}

function refresh() {
  renderUserList();
  renderPreview();
}

function download(filename, text, type) {
  const blob = new Blob([text], { type });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 1000);
}

async function copyHtml() {
  const html = $('html-output').value;
  try {
    await navigator.clipboard.write([new ClipboardItem({
      'text/html': new Blob([html], { type: 'text/html' }),
      'text/plain': new Blob([html], { type: 'text/plain' })
    })]);
    setStatus('Signature copied. Paste it into Outlook > Settings > Signatures.', 'ok');
  } catch (e) {
    $('html-output').select();
    setStatus('Could not access the clipboard; the HTML is selected below, press Ctrl+C.', 'warn');
  }
}

function exportAll() {
  const signatures = state.users
    .filter((u) => u.mail)
    .map((u) => ({
      mail: u.mail,
      displayName: u.displayName,
      html: SignatureRenderer.renderSignature(state.template, u, { variant: 'new' }),
      replyHtml: SignatureRenderer.renderSignature(state.template, u, { variant: 'reply' })
    }));
  download('signatures-export.json', JSON.stringify({ generatedAt: new Date().toISOString(), signatures }, null, 2), 'application/json');
  setStatus('Exported ' + signatures.length + ' signatures for Deploy-Signatures.ps1.', 'ok');
}

function csvCell(value) {
  return '"' + String(value || '').replace(/"/g, '""') + '"';
}

function exportDataQuality() {
  const rows = [['Name', 'Email', 'Department', 'Missing fields']];
  for (const u of state.users) {
    const missing = SignatureRenderer.missingFields(state.template, u);
    if (missing.length) rows.push([u.displayName, u.mail, u.department, missing.join('; ')]);
  }
  download('signature-data-gaps.csv', rows.map((r) => r.map(csvCell).join(',')).join('\r\n'), 'text/csv');
  setStatus((rows.length - 1) + ' people have empty fields used by the signature.', rows.length > 1 ? 'warn' : 'ok');
}

async function graphGetAll(url, token) {
  const all = [];
  let next = url;
  while (next) {
    const res = await fetch(next, { headers: { Authorization: 'Bearer ' + token } });
    if (!res.ok) throw new Error('Microsoft Graph returned HTTP ' + res.status + ': ' + (await res.text()));
    const body = await res.json();
    all.push(...body.value);
    next = body['@odata.nextLink'];
  }
  return all;
}

async function getToken() {
  try {
    return (await pca.acquireTokenSilent({ scopes: SCOPES, account: state.account })).accessToken;
  } catch (e) {
    return (await pca.acquireTokenPopup({ scopes: SCOPES, account: state.account })).accessToken;
  }
}

async function loadDirectory() {
  setStatus('Loading staff from Microsoft 365...');
  const token = await getToken();
  const users = (await graphGetAll(GRAPH_USERS, token))
    .filter((u) => u.mail)
    .sort((a, b) => (a.displayName || '').localeCompare(b.displayName || ''));
  state.users = users;
  state.demo = false;
  state.selectedId = users.length ? users[0].id : null;
  refresh();
  setStatus('Loaded ' + users.length + ' mailboxes from Microsoft 365.', 'ok');
}

async function signIn() {
  try {
    const result = await pca.loginPopup({ scopes: SCOPES });
    state.account = result.account;
    pca.setActiveAccount(result.account);
    $('signin').textContent = 'Signed in as ' + result.account.username;
    $('signin').disabled = true;
    $('reload').hidden = false;
    await loadDirectory();
  } catch (e) {
    setStatus('Sign-in failed: ' + e.message, 'error');
  }
}

async function init() {
  try {
    const res = await fetch('template.json', { cache: 'no-store' });
    state.template = await res.json();
  } catch (e) {
    setStatus('Could not load template.json: ' + e.message, 'error');
    return;
  }
  syncJson();
  renderBrandForm();
  $('preview-date').value = state.previewDate;
  refresh();

  $('search').addEventListener('input', (e) => { state.filter = e.target.value; renderUserList(); });
  $('variant').addEventListener('change', (e) => { state.variant = e.target.value; renderPreview(); });
  $('preview-date').addEventListener('change', (e) => { state.previewDate = e.target.value; renderPreview(); });
  $('template-json').addEventListener('input', (e) => {
    try {
      state.template = JSON.parse(e.target.value);
      renderBrandForm();
      refresh();
    } catch (err) {
      $('template-error').textContent = 'Invalid JSON: ' + err.message;
    }
  });
  $('copy').addEventListener('click', copyHtml);
  $('download-template').addEventListener('click', () => {
    download('template.json', JSON.stringify(state.template, null, 2) + '\n', 'application/json');
    setStatus('Commit the downloaded template.json to the signatures folder to roll it out to everyone.', 'ok');
  });
  $('export-all').addEventListener('click', exportAll);
  $('export-gaps').addEventListener('click', exportDataQuality);
  $('reload').addEventListener('click', () => loadDirectory().catch((e) => setStatus(e.message, 'error')));

  if (/^0{8}-/.test(config.clientId)) {
    $('signin').disabled = true;
    setStatus('Demo mode with sample people. Set clientId in src/config.js and rebuild to connect to Microsoft 365.', 'warn');
    return;
  }
  pca = new PublicClientApplication({
    auth: {
      clientId: config.clientId,
      authority: 'https://login.microsoftonline.com/' + config.tenantId,
      redirectUri: config.baseUrl + '/auth-redirect.html'
    },
    cache: { cacheLocation: 'sessionStorage' }
  });
  await pca.initialize();
  $('signin').addEventListener('click', signIn);
  const account = pca.getAllAccounts()[0];
  if (account) {
    state.account = account;
    $('signin').textContent = 'Signed in as ' + account.username;
    $('signin').disabled = true;
    $('reload').hidden = false;
    loadDirectory().catch((e) => setStatus(e.message, 'error'));
  } else {
    setStatus('Showing sample people. Sign in with a Microsoft 365 admin account to load your staff.');
  }
}

init();
