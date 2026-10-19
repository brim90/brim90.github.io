// Tenant-specific settings for the Outlook add-in. Edit, then run `npm run build`.
module.exports = {
  // Application (client) ID of the Entra ID app registration (see README).
  clientId: '00000000-0000-0000-0000-000000000000',
  // Directory (tenant) ID, or your tenant domain e.g. 'waterbabies.co.uk'.
  tenantId: 'organizations',
  // Absolute URL this folder is published at (classic Outlook gives the
  // add-in no page URL, so everything it fetches must be absolute). Also used
  // to generate addin/manifest.xml. Keep it on a Water Babies domain: images
  // and links on a different domain from the sender look like phishing to
  // spam filters.
  baseUrl: 'https://signatures.waterbabies.co.uk',
  // How long a profile and the designs are reused before being re-fetched.
  cacheMinutes: 60
};
