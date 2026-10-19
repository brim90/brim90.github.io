// Tenant-specific settings. Edit these, then run `npm run build`.
module.exports = {
  // Application (client) ID of the Entra ID app registration (see README).
  clientId: '00000000-0000-0000-0000-000000000000',
  // Directory (tenant) ID, or your tenant domain e.g. 'waterbabies.co.uk'.
  tenantId: 'organizations',
  // Absolute URL this folder is published at. The Outlook add-in runs without
  // a page URL in classic Outlook, so everything it fetches must be absolute.
  baseUrl: 'https://danbrimmicombe.co.uk/signatures',
  // How long the add-in reuses a cached profile/template before refreshing.
  cacheMinutes: 60
};
