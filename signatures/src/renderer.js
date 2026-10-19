/*
 * Signature renderer shared by the admin console, the Outlook add-in and the
 * tests. It has no dependencies and no DOM access so it also runs inside the
 * JavaScript-only runtime that classic Outlook on Windows uses for
 * event-based add-ins.
 *
 * Template syntax (a small Mustache subset):
 *   {{path.to.value}}        HTML-escaped value
 *   {{{path.to.value}}}      raw value (only use for trusted template config)
 *   {{#path}}...{{/path}}    section: rendered when the value is truthy; for
 *                            arrays it repeats once per item, with the item
 *                            as the current scope
 *   {{^path}}...{{/path}}    inverted section: rendered when the value is falsy
 *   {{.}}                    the current item inside an array section
 */

const PROFILE_FIELDS = [
  'id', 'displayName', 'givenName', 'surname', 'jobTitle', 'department',
  'companyName', 'officeLocation', 'businessPhones', 'mobilePhone', 'mail',
  'userPrincipalName', 'streetAddress', 'city', 'state', 'postalCode',
  'country', 'employeeType', 'onPremisesExtensionAttributes'
];

// $select used for every Graph call so the admin console, the add-in and the
// PowerShell script all see the same fields.
const GRAPH_SELECT = PROFILE_FIELDS.join(',');

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function lookup(stack, path) {
  if (path === '.') return stack[stack.length - 1];
  const parts = path.split('.');
  for (let i = stack.length - 1; i >= 0; i--) {
    const scope = stack[i];
    if (scope !== null && typeof scope === 'object' && parts[0] in scope) {
      let value = scope;
      for (const part of parts) {
        if (value === null || value === undefined) return undefined;
        value = value[part];
      }
      return value;
    }
  }
  return undefined;
}

function isTruthy(value) {
  if (Array.isArray(value)) return value.length > 0;
  if (typeof value === 'string') return value.trim().length > 0;
  return Boolean(value);
}

function parse(template) {
  const re = /\{\{\{\s*([\w.]+)\s*\}\}\}|\{\{\s*([#^/]?)\s*([\w.]+)\s*\}\}/g;
  const root = { children: [] };
  const stack = [root];
  let last = 0;
  let match;
  while ((match = re.exec(template)) !== null) {
    const current = stack[stack.length - 1];
    if (match.index > last) current.children.push({ type: 'text', value: template.slice(last, match.index) });
    last = re.lastIndex;
    if (match[1]) {
      current.children.push({ type: 'raw', path: match[1] });
      continue;
    }
    const sigil = match[2];
    const path = match[3];
    if (sigil === '#' || sigil === '^') {
      const node = { type: sigil === '#' ? 'section' : 'inverted', path, children: [] };
      current.children.push(node);
      stack.push(node);
    } else if (sigil === '/') {
      if (stack.length === 1 || current.path !== path) {
        throw new Error(`Unexpected closing tag {{/${path}}}`);
      }
      stack.pop();
    } else {
      current.children.push({ type: 'var', path });
    }
  }
  if (stack.length > 1) throw new Error(`Unclosed section {{#${stack[stack.length - 1].path}}}`);
  if (last < template.length) root.children.push({ type: 'text', value: template.slice(last) });
  return root;
}

function renderNodes(nodes, stack) {
  let out = '';
  for (const node of nodes) {
    if (node.type === 'text') {
      out += node.value;
    } else if (node.type === 'var' || node.type === 'raw') {
      const value = lookup(stack, node.path);
      if (value === undefined || value === null) continue;
      out += node.type === 'var' ? escapeHtml(value) : String(value);
    } else if (node.type === 'section') {
      const value = lookup(stack, node.path);
      if (!isTruthy(value)) continue;
      if (Array.isArray(value)) {
        for (const item of value) out += renderNodes(node.children, stack.concat([item]));
      } else {
        out += renderNodes(node.children, typeof value === 'object' ? stack.concat([value]) : stack);
      }
    } else if (node.type === 'inverted') {
      if (!isTruthy(lookup(stack, node.path))) out += renderNodes(node.children, stack);
    }
  }
  return out;
}

function renderString(template, context) {
  return renderNodes(parse(template).children, [context]);
}

function telHref(number) {
  if (!number) return '';
  const trimmed = String(number).trim();
  const digits = trimmed.replace(/[^\d]/g, '');
  return digits ? 'tel:' + (trimmed.startsWith('+') ? '+' : '') + digits : '';
}

function clean(value) {
  if (value === undefined || value === null) return '';
  return String(value).trim();
}

/**
 * Turns a raw Microsoft Graph user object into the `user` object templates see.
 * `fieldMap` lets the template pull custom fields (pronouns, a booking link,
 * a personal LinkedIn URL...) from Entra ID extension attributes, e.g.
 *   { "pronouns": "onPremisesExtensionAttributes.extensionAttribute1" }
 */
function buildUser(profile, template) {
  const p = profile || {};
  const defaults = (template && template.defaults) || {};
  const fieldMap = (template && template.fieldMap) || {};
  const businessPhones = Array.isArray(p.businessPhones) ? p.businessPhones.filter(Boolean) : [];

  const user = {
    displayName: clean(p.displayName),
    givenName: clean(p.givenName),
    surname: clean(p.surname),
    jobTitle: clean(p.jobTitle),
    department: clean(p.department),
    companyName: clean(p.companyName),
    officeLocation: clean(p.officeLocation),
    phone: clean(businessPhones[0]),
    mobilePhone: clean(p.mobilePhone),
    email: clean(p.mail || p.userPrincipalName),
    streetAddress: clean(p.streetAddress),
    city: clean(p.city),
    state: clean(p.state),
    postalCode: clean(p.postalCode),
    country: clean(p.country)
  };

  for (const [name, path] of Object.entries(fieldMap)) {
    const value = path.split('.').reduce((acc, key) => (acc == null ? undefined : acc[key]), p);
    user[name] = clean(value);
  }

  // Organisation-wide fallbacks, e.g. the main switchboard number.
  for (const [name, value] of Object.entries(defaults)) {
    if (!user[name]) user[name] = clean(value);
  }

  if (!user.displayName) user.displayName = [user.givenName, user.surname].filter(Boolean).join(' ');
  user.firstName = user.givenName || user.displayName.split(' ')[0] || '';
  user.phoneHref = telHref(user.phone);
  user.mobilePhoneHref = telHref(user.mobilePhone);
  user.emailHref = user.email ? 'mailto:' + user.email : '';
  user.address = [user.streetAddress, user.city, user.postalCode].filter(Boolean).join(', ');
  user.titleLine = [user.jobTitle, user.department].filter(Boolean).join(' | ');
  return user;
}

function campaignApplies(campaign, user, now) {
  if (!campaign || campaign.enabled === false) return false;
  const time = now.getTime();
  if (campaign.start && time < Date.parse(campaign.start)) return false;
  // An end date with no time means "to the end of that day".
  if (campaign.end) {
    const end = /^\d{4}-\d{2}-\d{2}$/.test(campaign.end) ? Date.parse(campaign.end + 'T23:59:59') : Date.parse(campaign.end);
    if (time > end) return false;
  }
  const departments = (campaign.departments || []).map((d) => d.toLowerCase());
  if (departments.length && !departments.includes(user.department.toLowerCase())) return false;
  const exclude = (campaign.excludeEmails || []).map((e) => e.toLowerCase());
  if (exclude.includes(user.email.toLowerCase())) return false;
  return true;
}

/**
 * Picks the first campaign banner that is live for this user and date.
 * Campaigns are listed in priority order in template.json.
 */
function selectCampaign(template, user, now) {
  const campaigns = (template && template.campaigns) || [];
  return campaigns.find((c) => campaignApplies(c, user, now)) || null;
}

function buildContext(template, profile, options) {
  const opts = options || {};
  const now = opts.now || new Date();
  const user = buildUser(profile, template);
  return {
    user,
    brand: (template && template.brand) || {},
    campaign: selectCampaign(template, user, now),
    year: now.getFullYear(),
    isReply: opts.variant === 'reply'
  };
}

/**
 * Renders the signature HTML for one person.
 *   variant: "new" (full signature) or "reply" (the shorter reply/forward
 *   signature; falls back to the full one if the template has none).
 */
function renderSignature(template, profile, options) {
  const opts = options || {};
  const context = buildContext(template, profile, opts);
  const html = opts.variant === 'reply' && template.replyHtml ? template.replyHtml : template.html;
  return renderString(html, context).replace(/>\s+</g, '><').trim();
}

/**
 * Lists the fields a template actually uses that are empty for this person,
 * so admins can fix the source data in Entra ID before rollout.
 */
function missingFields(template, profile) {
  const user = buildUser(profile, template);
  const source = (template.html || '') + (template.replyHtml || '');
  const used = new Set();
  const re = /\{\{\{?\s*[#^]?\s*user\.(\w+)/g;
  let match;
  while ((match = re.exec(source)) !== null) used.add(match[1]);
  const derived = { phoneHref: 'phone', mobilePhoneHref: 'mobilePhone', emailHref: 'email', titleLine: 'jobTitle', address: 'streetAddress', firstName: 'givenName' };
  const missing = new Set();
  for (const field of used) {
    const base = derived[field] || field;
    if (!user[base]) missing.add(base);
  }
  return Array.from(missing);
}

const SignatureRenderer = {
  GRAPH_SELECT,
  PROFILE_FIELDS,
  escapeHtml,
  renderString,
  buildUser,
  buildContext,
  selectCampaign,
  renderSignature,
  missingFields,
  telHref
};

module.exports = SignatureRenderer;
