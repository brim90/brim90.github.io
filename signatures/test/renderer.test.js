const test = require('node:test');
const assert = require('node:assert');
const fs = require('node:fs');
const path = require('node:path');
const R = require('../src/renderer');

const template = JSON.parse(fs.readFileSync(path.join(__dirname, '..', 'template.json'), 'utf8'));
const addressHtml = R.escapeHtml(template.brand.address);

const full = {
  displayName: 'Alex Morgan', givenName: 'Alex', jobTitle: 'Head of Swim Programmes', department: 'Operations',
  businessPhones: ['0330 223 6700'], mobilePhone: '07700 900123', mail: 'alex.morgan@waterbabies.co.uk'
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
  for (const text of ['Alex Morgan', 'Head of Swim Programmes', '0330 223 6700', '07700 900123',
    'href="tel:03302236700"', 'href="tel:07700900123"', addressHtml, 'Water Babies Group Limited']) {
    assert.ok(html.includes(text), 'missing ' + text);
  }
  assert.ok(!html.includes('{{'), 'unrendered placeholder');
});

test('website is an underlined hyperlink', () => {
  const html = R.renderSignature(template, full);
  assert.match(html, /<a href="https:\/\/www\.waterbabies\.co\.uk" style="[^"]*text-decoration:underline;[^"]*">waterbabies\.co\.uk<\/a>/);
});

test('people without their own number get the switchboard', () => {
  const html = R.renderSignature(template, { displayName: 'Sam Patel', mobilePhone: '07700 900456', mail: 'sam@waterbabies.co.uk' });
  assert.ok(html.includes('0330 223 6777'));
  assert.ok(html.includes('07700 900456'));
});

test('empty fields drop their whole line', () => {
  const t = { ...template, defaults: {} };
  const profile = { displayName: 'Jordan Lee', mail: 'jordan@waterbabies.co.uk' };
  const html = R.renderSignature(t, profile);
  assert.ok(!html.includes('T:</span>'), 'office label shown without a number');
  assert.ok(!html.includes('M:</span>'), 'mobile label shown without a mobile');
  assert.deepStrictEqual(R.missingFields(t, profile).sort(), ['jobTitle', 'mobilePhone', 'phone']);
});

test('profile values cannot inject HTML', () => {
  const html = R.renderSignature(template, { ...full, jobTitle: '<script>alert(1)</script>' });
  assert.ok(!html.includes('<script>'));
  assert.ok(html.includes('&lt;script&gt;'));
});

test('reply variant uses the short template, preferring the mobile', () => {
  const html = R.renderSignature(template, full, { variant: 'reply' });
  assert.ok(html.includes('Alex Morgan'));
  assert.ok(html.includes('07700 900123'));
  assert.ok(!html.includes('0330 223 6700'));
  assert.ok(!html.includes(addressHtml));
  const noMobile = R.renderSignature(template, { ...full, mobilePhone: '' }, { variant: 'reply' });
  assert.ok(noMobile.includes('0330 223 6700'));
});

test('campaign banners respect dates, departments and enabled flag', () => {
  const banner = { name: 'Appeal', imageUrl: 'https://x/b.png', link: 'https://x', start: '2026-11-01', end: '2026-11-30', departments: ['Operations'] };
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

test('custom fields can be mapped from extension attributes', () => {
  const t = { ...template, fieldMap: { pronouns: 'onPremisesExtensionAttributes.extensionAttribute1' }, html: '{{#user.pronouns}}({{user.pronouns}}){{/user.pronouns}}' };
  assert.strictEqual(R.renderSignature(t, { onPremisesExtensionAttributes: { extensionAttribute1: 'she/her' } }), '(she/her)');
  assert.strictEqual(R.renderSignature(t, {}), '');
});
