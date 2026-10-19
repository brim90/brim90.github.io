const test = require('node:test');
const assert = require('node:assert');
const fs = require('node:fs');
const path = require('node:path');
const R = require('../src/renderer');

const root = path.join(__dirname, '..');
const config = JSON.parse(fs.readFileSync(path.join(root, 'signatures.json'), 'utf8'));
const read = (file) => fs.readFileSync(path.join(root, file), 'utf8');
const designs = Object.fromEntries(Object.entries(config.designs)
  .map(([name, files]) => [name, { new: read(files.new), reply: files.reply && read(files.reply) }]));
const ASSETS = 'https://example.org/signatures/assets';
const render = (profile, variant) => R.renderSignature(config, designs, profile, ASSETS, variant).html;

const dan = { displayName: 'Dan Brimmicombe', jobTitle: 'Associate Director of Technology', mail: 'dan@waterbabies.co.uk' };
const james = { displayName: 'James Brookes', jobTitle: 'Managing Director, UK & EU', mobilePhone: '07700 900456', mail: 'james@waterbabies.co.uk' };

test('values are escaped and sections follow emptiness', () => {
  assert.strictEqual(R.renderString('{{a}}{{#b}}B{{/b}}{{^b}}no b{{/b}}', { a: '<i>&', b: ' ' }), '&lt;i&gt;&amp;no b');
});

test('broken designs fail loudly', () => {
  assert.throws(() => R.renderString('{{#a}}', {}), /never closed/);
  assert.throws(() => R.renderString('{{#a}}{{/b}}', {}), /Unexpected/);
});

test('every design renders with no leftover placeholders or comments', () => {
  for (const [name, files] of Object.entries(designs)) {
    for (const [variant, template] of Object.entries(files)) {
      if (!template) continue;
      const html = R.renderString(template, { user: R.buildUser(james, config.company), company: config.company, assets: ASSETS });
      assert.ok(!html.includes('{{'), `${name} (${variant}) has an unrendered placeholder`);
      assert.ok(!html.includes('<!--'), `${name} (${variant}) still contains a comment`);
    }
  }
});

test('both versions carry the person and the legal details', () => {
  for (const variant of ['new', 'reply']) {
    const html = render(dan, variant);
    for (const text of ['Dan Brimmicombe', 'Associate Director of Technology',
      'Water Babies Ltd. Registered in England &amp; Wales No. 04322357',
      'Registered office: Winslade House, Winslade Park, Manor Drive,', 'Clyst St Mary, Exeter, EX5 1FY']) {
      assert.ok(html.includes(text), `${variant} is missing ${text}`);
    }
  }
});

test('switchboard is used when there is no office number; mobile only when set', () => {
  const html = render(dan);
  assert.ok(html.includes('href="tel:03302236777"'));
  assert.ok(!html.includes('>M</span>'));
  const withMobile = render(james);
  assert.ok(withMobile.includes('href="tel:07700900456"'));
  assert.ok(withMobile.indexOf('>M</span>') < withMobile.indexOf('>T</span>'), 'mobile comes first, as in the guide');
  assert.ok(render({ ...dan, businessPhones: ['01392 000000'] }).includes('01392 000000'));
});

test('website is an underlined link', () => {
  assert.match(render(dan), /<a href="https:\/\/www\.waterbabies\.co\.uk" target="_blank" style="color:#FF6600;text-decoration:underline;">/);
});

test('reply signature is short: no images, mobile preferred, legal line kept', () => {
  const html = render(james, 'reply');
  assert.ok(!html.includes('<img'));
  assert.ok(html.includes('James Brookes') && html.includes('Managing Director, UK &amp; EU'));
  assert.ok(html.includes('07700 900456') && !html.includes('0330 223 6777'));
  assert.ok(render(dan, 'reply').includes('href="tel:03302236777"'), 'switchboard when there is no mobile');
  assert.match(html, /text-decoration:underline;"><span style="color:#FF6600;">waterbabies\.co\.uk/);
});

test('designs without a reply version use the full one', () => {
  const only = { standard: { new: designs.standard.new } };
  assert.strictEqual(R.renderSignature(config, only, dan, ASSETS, 'reply').html, render(dan, 'new'));
});

test('every image has absolute src, width and height (needed by Outlook desktop)', () => {
  const imgs = render(dan).match(/<img [^>]*>/g);
  assert.strictEqual(imgs.length, 3, 'keep the image count low: logo, dark-mode logo, icon strip');
  for (const img of imgs) {
    assert.match(img, new RegExp('src="' + ASSETS + '/[\\w-]+\\.png"'));
    assert.match(img, / width="\d+"/);
    assert.match(img, / height="\d+"/);
  }
  for (const file of render(dan).match(/assets\/[\w-]+\.png/g)) {
    assert.ok(fs.existsSync(path.join(root, file)), file + ' is missing');
  }
});

test('profile values cannot inject HTML', () => {
  const html = render({ ...dan, jobTitle: '<script>x</script>' });
  assert.ok(!html.includes('<script>') && html.includes('&lt;script&gt;'));
});

test('rules pick a design, first match wins, otherwise the default', () => {
  const cfg = { ...config, rules: [{ design: 'other', when: { department: ['Franchise Support'] } }] };
  assert.strictEqual(R.chooseDesign(cfg, R.buildUser({ department: 'franchise support' })), 'other');
  assert.strictEqual(R.chooseDesign(cfg, R.buildUser({ department: 'Finance' })), config.defaultDesign);
  assert.throws(() => R.renderSignature(cfg, designs, { department: 'Franchise Support' }, ASSETS), /not listed/);
});

test('white logo is hidden unless the app supports dark mode', () => {
  const html = render(dan);
  const dark = html.match(/<img class="wb-logo-dark"[^>]*>/)[0];
  assert.match(dark, /display:none;/);
  assert.match(dark, /mso-hide:all;/, 'classic Outlook must never show it');
  assert.match(dark, /logo-white\.png/);
  assert.match(html, /<img class="wb-logo-light"[^>]*style="display:block;/);
  assert.match(html, /@media \(prefers-color-scheme: dark\)/);
});
