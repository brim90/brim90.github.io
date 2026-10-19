// Loaded by auth-redirect.html: hands the sign-in result back to the admin console window.
const { broadcastResponseToMainFrame } = require('@azure/msal-browser/redirect-bridge');

broadcastResponseToMainFrame().catch((e) => {
  document.body.textContent = 'Sign-in could not be completed: ' + e.message;
});
