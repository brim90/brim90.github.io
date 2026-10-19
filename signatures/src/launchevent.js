/*
 * Outlook add-in: inserts the Water Babies signature whenever someone starts
 * a new email, reply or forward, filled in from their Microsoft 365 profile.
 *
 * Runs in commands.html (Outlook on the web, new Outlook, Outlook for Mac,
 * Outlook for iOS and Android) and as a plain script in classic Outlook for
 * Windows, which has no DOM, no page URL and no localStorage.
 */
/* global Office, OfficeRuntime */
const { createNestablePublicClientApplication } = require('@azure/msal-browser');
const SignatureRenderer = require('./renderer');
const config = require('./config');

const GRAPH_ME = 'https://graph.microsoft.com/v1.0/me?$select=' + SignatureRenderer.GRAPH_SELECT;
const NETWORK_TIMEOUT_MS = 8000;

const storage = {
  async get(key) {
    try {
      if (typeof OfficeRuntime !== 'undefined' && OfficeRuntime.storage) {
        const raw = await OfficeRuntime.storage.getItem(key);
        return raw ? JSON.parse(raw) : null;
      }
      if (typeof localStorage !== 'undefined') {
        const raw = localStorage.getItem(key);
        return raw ? JSON.parse(raw) : null;
      }
    } catch (e) {
      console.warn('Signature cache read failed', e);
    }
    return null;
  },
  async set(key, value) {
    try {
      const raw = JSON.stringify(value);
      if (typeof OfficeRuntime !== 'undefined' && OfficeRuntime.storage) {
        await OfficeRuntime.storage.setItem(key, raw);
      } else if (typeof localStorage !== 'undefined') {
        localStorage.setItem(key, raw);
      }
    } catch (e) {
      console.warn('Signature cache write failed', e);
    }
  }
};

function withTimeout(promise, ms) {
  return Promise.race([
    promise,
    new Promise((_, reject) => setTimeout(() => reject(new Error('Timed out after ' + ms + 'ms')), ms))
  ]);
}

// Returns a cached copy younger than cacheMinutes, otherwise loads a fresh
// one, falling back to the stale copy if loading fails (e.g. offline).
async function cached(key, load) {
  const entry = await storage.get(key);
  if (entry && Date.now() - entry.savedAt < config.cacheMinutes * 60 * 1000) return entry.value;
  try {
    const value = await withTimeout(load(), NETWORK_TIMEOUT_MS);
    await storage.set(key, { savedAt: Date.now(), value });
    return value;
  } catch (e) {
    console.warn('Refreshing ' + key + ' failed, using cached copy', e);
    if (entry) return entry.value;
    throw e;
  }
}

async function fetchText(path) {
  const res = await fetch(config.baseUrl + '/' + path + '?t=' + Date.now(), { cache: 'no-store' });
  if (!res.ok) throw new Error('Downloading ' + path + ' failed: HTTP ' + res.status);
  return res.text();
}

// signatures.json plus the HTML of every design it lists.
function loadDesigns() {
  return cached('signature.designs', async () => {
    const signatureConfig = JSON.parse(await fetchText('signatures.json'));
    const designs = {};
    for (const [name, path] of Object.entries(signatureConfig.designs)) designs[name] = await fetchText(path);
    return { config: signatureConfig, designs };
  });
}

let pcaPromise;
async function fetchGraphProfile() {
  if (!pcaPromise) {
    pcaPromise = createNestablePublicClientApplication({
      auth: { clientId: config.clientId, authority: 'https://login.microsoftonline.com/' + config.tenantId }
    });
  }
  const pca = await pcaPromise;
  const request = { scopes: ['User.Read'] };
  const account = pca.getActiveAccount() || pca.getAllAccounts()[0];
  if (account) request.account = account;
  // Add-ins can't show a sign-in prompt here, so this relies on admin consent.
  const token = await pca.acquireTokenSilent(request);
  const res = await fetch(GRAPH_ME, { headers: { Authorization: 'Bearer ' + token.accessToken } });
  if (!res.ok) throw new Error('Microsoft Graph /me failed: HTTP ' + res.status);
  return res.json();
}

async function loadProfile() {
  const mailboxUser = Office.context.mailbox.userProfile;
  try {
    return await cached('signature.profile.' + mailboxUser.emailAddress.toLowerCase(), fetchGraphProfile);
  } catch (e) {
    // Last resort: the two things Outlook always knows about the sender.
    console.warn('Using basic mailbox profile for signature', e);
    return { displayName: mailboxUser.displayName, mail: mailboxUser.emailAddress };
  }
}

function officeCall(fn) {
  return new Promise((resolve, reject) => {
    fn((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
      else reject(result.error);
    });
  });
}

async function applySignature() {
  const item = Office.context.mailbox.item;
  const [{ config: signatureConfig, designs }, profile] = await Promise.all([loadDesigns(), loadProfile()]);
  const { html } = SignatureRenderer.renderSignature(signatureConfig, designs, profile, config.baseUrl + '/assets');
  // Stop Outlook adding the person's own signature as well as ours.
  if (item.disableClientSignatureAsync) {
    await officeCall((cb) => item.disableClientSignatureAsync(cb)).catch((e) => console.warn(e));
  }
  await officeCall((cb) => item.body.setSignatureAsync(html, { coercionType: Office.CoercionType.Html }, cb));
}

function onNewMessageComposeHandler(event) {
  applySignature()
    .catch((e) => console.error('Signature could not be applied', e))
    .finally(() => event.completed());
}

Office.onReady(() => {});
Office.actions.associate('onNewMessageComposeHandler', onNewMessageComposeHandler);
