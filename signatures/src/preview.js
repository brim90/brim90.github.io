// Preview page: renders each design for whatever details are typed in, so a
// design change can be checked before it's committed.
const SignatureRenderer = require('./renderer');

const $ = (id) => document.getElementById(id);
let signatureConfig;
const designs = {};

function profileFromForm() {
  return {
    displayName: $('name').value,
    jobTitle: $('title').value,
    businessPhones: [$('phone').value],
    mobilePhone: $('mobile').value,
    mail: 'preview@waterbabies.co.uk'
  };
}

function render() {
  const design = designs[$('design').value];
  const template = ($('variant').value === 'reply' && design.reply) || design.new;
  const html = SignatureRenderer.renderString(template, {
    user: SignatureRenderer.buildUser(profileFromForm(), signatureConfig.company),
    company: signatureConfig.company,
    assets: new URL('assets', location.href).href
  });
  $('preview').srcdoc = '<!doctype html><html><head><meta charset="utf-8"></head><body style="margin:16px;background:#fff;">' +
    '<p style="font-family:Arial,sans-serif;font-size:10pt;color:#333;margin:0 0 16px;">Kind regards,</p>' + html + '</body></html>';
  $('html').value = html;
}

async function copySignature() {
  const html = $('html').value;
  try {
    await navigator.clipboard.write([new ClipboardItem({
      'text/html': new Blob([html], { type: 'text/html' }),
      'text/plain': new Blob([html], { type: 'text/plain' })
    })]);
    $('status').textContent = 'Copied. Paste into a test email in Outlook to check it.';
  } catch (e) {
    $('html').select();
    $('status').textContent = 'Clipboard blocked: the HTML is selected below, press Ctrl+C.';
  }
}

async function init() {
  signatureConfig = await (await fetch('signatures.json', { cache: 'no-store' })).json();
  const load = async (path) => (await fetch(path, { cache: 'no-store' })).text();
  for (const [name, files] of Object.entries(signatureConfig.designs)) {
    designs[name] = { new: await load(files.new), reply: files.reply ? await load(files.reply) : null };
    $('design').append(new Option(name, name, name === signatureConfig.defaultDesign, name === signatureConfig.defaultDesign));
  }
  for (const id of ['design', 'variant', 'name', 'title', 'phone', 'mobile']) $(id).addEventListener('input', render);
  $('copy').addEventListener('click', copySignature);
  render();
}

init().catch((e) => { $('status').textContent = 'Could not load the designs: ' + e.message; });
