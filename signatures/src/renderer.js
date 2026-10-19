/*
 * Turns a design (an HTML file in designs/) plus a person's Microsoft 365
 * profile into their signature. Shared by the Outlook add-in, the preview
 * page and the tests. No dependencies and no DOM, so it also runs in classic
 * Outlook's JavaScript-only runtime.
 *
 * Design syntax (a small Mustache subset):
 *   {{user.jobTitle}}              value, HTML-escaped
 *   {{#user.mobilePhone}}…{{/user.mobilePhone}}
 *                                  only included when the value isn't empty
 *   {{^user.mobilePhone}}…{{/user.mobilePhone}}
 *                                  only included when the value is empty
 */

// Profile fields read from Microsoft Graph.
const GRAPH_SELECT = [
  'displayName', 'givenName', 'surname', 'jobTitle', 'department', 'companyName',
  'officeLocation', 'businessPhones', 'mobilePhone', 'mail', 'userPrincipalName'
].join(',');

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function lookup(context, path) {
  return path.split('.').reduce((value, key) => (value == null ? undefined : value[key]), context);
}

function isEmpty(value) {
  return value == null || value === false || String(value).trim() === '';
}

function parse(template) {
  const re = /\{\{\s*([#^/]?)\s*([\w.]+)\s*\}\}/g;
  const root = { children: [] };
  const stack = [root];
  let last = 0;
  let match;
  while ((match = re.exec(template)) !== null) {
    const current = stack[stack.length - 1];
    if (match.index > last) current.children.push({ text: template.slice(last, match.index) });
    last = re.lastIndex;
    const [, sigil, path] = match;
    if (sigil === '#' || sigil === '^') {
      const node = { path, inverted: sigil === '^', children: [] };
      current.children.push(node);
      stack.push(node);
    } else if (sigil === '/') {
      if (stack.length === 1 || current.path !== path) throw new Error(`Unexpected {{/${path}}}`);
      stack.pop();
    } else {
      current.children.push({ variable: path });
    }
  }
  if (stack.length > 1) throw new Error(`{{#${stack[stack.length - 1].path}}} is never closed`);
  if (last < template.length) root.children.push({ text: template.slice(last) });
  return root;
}

function renderNodes(nodes, context) {
  let out = '';
  for (const node of nodes) {
    if (node.text !== undefined) {
      out += node.text;
    } else if (node.variable) {
      const value = lookup(context, node.variable);
      if (!isEmpty(value)) out += escapeHtml(value);
    } else if (isEmpty(lookup(context, node.path)) === node.inverted) {
      out += renderNodes(node.children, context);
    }
  }
  return out;
}

function renderString(template, context) {
  const withoutComments = template.replace(/<!--(?!\[if)[\s\S]*?-->/g, '');
  return renderNodes(parse(withoutComments).children, context).replace(/>\s+</g, '><').trim();
}

function telHref(number) {
  const trimmed = String(number || '').trim();
  const digits = trimmed.replace(/\D/g, '');
  return digits ? 'tel:' + (trimmed.startsWith('+') ? '+' : '') + digits : '';
}

function clean(value) {
  return value == null ? '' : String(value).trim();
}

// The `user` object designs see. People without their own office number get
// the company switchboard.
function buildUser(profile, company) {
  const p = profile || {};
  const businessPhones = Array.isArray(p.businessPhones) ? p.businessPhones : [];
  const user = {
    displayName: clean(p.displayName) || [clean(p.givenName), clean(p.surname)].filter(Boolean).join(' '),
    jobTitle: clean(p.jobTitle),
    department: clean(p.department),
    companyName: clean(p.companyName),
    officeLocation: clean(p.officeLocation),
    phone: clean(businessPhones.find((n) => clean(n))) || clean(company && company.phone),
    mobilePhone: clean(p.mobilePhone),
    email: clean(p.mail || p.userPrincipalName)
  };
  user.phoneHref = telHref(user.phone);
  user.mobilePhoneHref = telHref(user.mobilePhone);
  return user;
}

/**
 * Picks the design for a person. Rules in signatures.json are checked in
 * order and the first match wins, e.g.
 *   { "design": "franchise", "when": { "department": ["Franchise Support"] } }
 * matches when every listed field equals one of its values (ignoring case).
 */
function chooseDesign(config, user) {
  for (const rule of config.rules || []) {
    const matches = Object.entries(rule.when || {}).every(([field, values]) =>
      [].concat(values).some((v) => String(v).toLowerCase() === clean(user[field]).toLowerCase()));
    if (matches) return rule.design;
  }
  return config.defaultDesign;
}

/**
 * config:  parsed signatures.json
 * designs: { designName: { new: html, reply: html } } (reply is optional)
 * assets:  URL of the assets folder the images are served from
 * variant: "new" for a new email, "reply" for replies and forwards
 */
function renderSignature(config, designs, profile, assets, variant) {
  const user = buildUser(profile, config.company);
  const design = chooseDesign(config, user);
  if (!designs[design]) throw new Error(`Design "${design}" is not listed in signatures.json`);
  const template = (variant === 'reply' && designs[design].reply) || designs[design].new;
  const html = renderString(template, { user, company: config.company, assets });
  return { design, html };
}

module.exports = { GRAPH_SELECT, escapeHtml, renderString, buildUser, chooseDesign, renderSignature, telHref };
