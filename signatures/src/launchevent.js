/*
 * Outlook event-based add-in: inserts the organisation's signature whenever a
 * message or meeting is composed, using the sender's Microsoft 365 profile.
 *
 * Runs in two places:
 *   - commands.html (Outlook on the web, new Outlook, Outlook on Mac)
 *   - directly as a JS file in classic Outlook on Windows, where there is no
 *     DOM, no window.location and no localStorage.
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

// Returns a fresh value from `load`, or the cached copy if loading fails.
// A cached copy younger than cacheMinutes is used without touching the network
// unless `refresh` is set.
async function cached(key, load, refresh) {
  const entry = await storage.get(key);
  const maxAge = config.cacheMinutes * 60 * 1000;
  if (!refresh && entry && Date.now() - entry.savedAt < maxAge) return entry.value;
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

async function loadTemplate(refresh) {
  return cached('signature.template', async () => {
    const res = await fetch(config.baseUrl + '/template.json?t=' + Date.now(), { cache: 'no-store' });
    if (!res.ok) throw new Error('Template download failed: HTTP ' + res.status);
    return res.json();
  }, refresh);
}

let pcaPromise;
function getPca() {
  if (!pcaPromise) {
    pcaPromise = createNestablePublicClientApplication({
      auth: {
        clientId: config.clientId,
        authority: 'https://login.microsoftonline.com/' + config.tenantId
      }
    });
  }
  return pcaPromise;
}

async function fetchGraphProfile() {
  const pca = await getPca();
  const request = { scopes: ['User.Read'] };
  const account = pca.getActiveAccount() || pca.getAllAccounts()[0];
  if (account) request.account = account;
  // Event-based add-ins cannot show sign-in UI, so only silent auth is possible.
  // Tenant-wide admin consent for User.Read makes this succeed for everyone.
  const token = await pca.acquireTokenSilent(request);
  const res = await fetch(GRAPH_ME, { headers: { Authorization: 'Bearer ' + token.accessToken } });
  if (!res.ok) throw new Error('Graph /me failed: HTTP ' + res.status);
  return res.json();
}

async function loadProfile(refresh) {
  const mailboxUser = Office.context.mailbox.userProfile;
  const cacheKey = 'signature.profile.' + mailboxUser.emailAddress.toLowerCase();
  try {
    return await cached(cacheKey, fetchGraphProfile, refresh);
  } catch (e) {
    // Last resort: the two fields Outlook always knows about the sender.
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

async function getVariant(item) {
  if (item.itemType === Office.MailboxEnums.ItemType.Appointment) return 'reply';
  try {
    const value = await officeCall((cb) => item.getComposeTypeAsync(cb));
    return value.composeType === 'newMail' ? 'new' : 'reply';
  } catch (e) {
    return 'new';
  }
}

async function applySignature(refresh) {
  const item = Office.context.mailbox.item;
  const [template, profile, variant] = await Promise.all([loadTemplate(refresh), loadProfile(refresh), getVariant(item)]);
  const html = SignatureRenderer.renderSignature(template, profile, { variant });
  // Stop Outlook adding the user's own signature on top of ours.
  if (item.disableClientSignatureAsync) {
    await officeCall((cb) => item.disableClientSignatureAsync(cb)).catch((e) => console.warn(e));
  }
  await officeCall((cb) => item.body.setSignatureAsync(html, { coercionType: Office.CoercionType.Html }, cb));
}

function makeHandler(refresh) {
  return (event) => {
    applySignature(refresh)
      .catch((e) => console.error('Signature could not be applied', e))
      .finally(() => event.completed());
  };
}

Office.onReady(() => {});
Office.actions.associate('onNewMessageComposeHandler', makeHandler(false));
Office.actions.associate('onNewAppointmentComposeHandler', makeHandler(false));
// Ribbon button: re-applies the signature with fresh data, e.g. after a
// profile change in Entra ID.
Office.actions.associate('insertSignature', makeHandler(true));
