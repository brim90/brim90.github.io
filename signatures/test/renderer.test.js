const test = require('node:test');
const assert = require('node:assert');
const fs = require('node:fs');
const path = require('node:path');
const R = require('../src/renderer');

const template = JSON.parse(fs.readFileSync(path.join(__dirname, '..', 'template.json'), 'utf8'));
const legalHtml = require('../src/renderer').escapeHtml(template.brand.legal);

const full = {
  displayName: 'Alex Morgan', givenName: 'Alex', jobTitle: 'Head of Fundraising', department: 'Fundraising',
  businessPhones: ['+44 20 7123 4567'], mobilePhone: '+44 7700 900123', mail: 'alex.morgan@childrensalliance.org.uk',
  streetAddress: '1 Example Street', city: 'London', postalCode: 'EC1A 1AA',
  onPremisesExtensionAttributes: { extensionAttribute1: 'they/them' }
};

test('variables are HTML-escaped and raw tags are not', () => {
  assert.strictEqual(R.renderString('{{a}}|{{{a}}}', { a: '<b>&' }), '&lt;b&gt;&amp;|<b>&');
});

test('sections hide empty values and loop over arrays', () => {
  const tpl = '{{#x}}X{{/x}}{{^x}}none{{/x}}{{#list}}[{{name}}:{{brand}}]{{/list}}';
  assert.strictEqual(R.renderString(tpl, { x: '  ', list: [{ name: 'a' }, { name: 'b' }], brand: 'B' }), 'none[a:B][b:B]');
});

test('unbalanced sections throw a helpful error', () => {
  assert.throws(() => R.renderString('{{#a}}', {}), /Unclosed section/);
  assert.throws(() => R.renderString('{{#a}}{{/b}}', {}), /Unexpected closing tag/);
});

test('full profile renders every personal field', () => {
  const html = R.renderSignature(template, full);
  for (const text of ['Alex Morgan', '(they/them)', 'Head of Fundraising', 'Fundraising', '+44 20 7123 4567',
    'href="tel:+442071234567"', 'href="tel:+447700900123"', 'mailto:alex.morgan@childrensalliance.org.uk',
    '1 Example Street, London, EC1A 1AA', 'LinkedIn']) {
    assert.ok(html.includes(text), 'missing ' + text);
  }
  assert.ok(!html.includes('{{'), 'unrendered placeholder');
});

test('empty fields drop their whole line', () => {
  const html = R.renderSignature(template, { displayName: 'Sam Patel', mail: 'sam@example.org' });
  assert.ok(!html.includes('>T<'), 'phone label shown without a phone');
  assert.ok(!html.includes('>M<'), 'mobile label shown without a mobile');
  assert.ok(!html.includes('()'), 'empty pronouns brackets');
  assert.deepStrictEqual(R.missingFields(template, { displayName: 'Sam Patel', mail: 'sam@example.org' }).sort(),
    ['department', 'jobTitle', 'mobilePhone', 'phone', 'pronouns', 'streetAddress']);
});

test('profile values cannot inject HTML', () => {
  const html = R.renderSignature(template, { ...full, jobTitle: '<script>alert(1)</script>' });
  assert.ok(!html.includes('<script>'));
  assert.ok(html.includes('&lt;script&gt;'));
});

test('reply variant uses the short template', () => {
  const html = R.renderSignature(template, full, { variant: 'reply' });
  assert.ok(html.includes('Alex Morgan'));
  assert.ok(!html.includes(legalHtml));
});

test('defaults fill gaps, e.g. a switchboard number', () => {
  const t = { ...template, defaults: { phone: '020 7000 0000' } };
  assert.ok(R.renderSignature(t, { displayName: 'A', mail: 'a@x.org' }).includes('020 7000 0000'));
  assert.ok(R.renderSignature(t, full).includes('+44 20 7123 4567'));
});

test('campaign banners respect dates, departments and enabled flag', () => {
  const banner = { name: 'Appeal', imageUrl: 'https://x/b.png', link: 'https://x', start: '2026-11-01', end: '2026-11-30', departments: ['Fundraising'] };
  const t = { ...template, campaigns: [banner] };
  const user = R.buildUser(full, t);
  assert.strictEqual(R.selectCampaign(t, user, new Date('2026-10-31T12:00:00')), null);
  assert.strictEqual(R.selectCampaign(t, user, new Date('2026-11-30T18:00:00')), banner);
  assert.strictEqual(R.selectCampaign(t, user, new Date('2026-12-01T09:00:00')), null);
  assert.strictEqual(R.selectCampaign(t, R.buildUser({ ...full, department: 'Finance' }, t), new Date('2026-11-15')), null);
  assert.strictEqual(R.selectCampaign({ campaigns: [{ ...banner, enabled: false }] }, user, new Date('2026-11-15')), null);
  assert.ok(R.renderSignature(t, full, { now: new Date('2026-11-15T12:00:00') }).includes('https://x/b.png'));
});

test('shipped template has no live campaign by default', () => {
  const html = R.renderSignature(template, full, { now: new Date('2026-11-15T12:00:00') });
  assert.ok(!html.includes('banner.png'));
});
