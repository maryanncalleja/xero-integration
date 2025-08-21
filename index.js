const express = require('express');
const { XeroClient } = require('xero-node');
const app = express();
const port = 5000;

// Your Xero client setup (replace with your credentials)
const xero = new XeroClient({
  clientId: '93B4125E75714BCCBB2B3CAB4E5AC7CF',
  clientSecret: 'm_2kdSW-3SkAzHYaIvPt8yhzSEkVk8myxfhIVrnAgzkBD9Za',
  redirectUris: ['http://localhost:5000/callback'],
  scopes: [
    'openid',
    'profile',
    'email',
    'offline_access',
    'accounting.transactions',
    'accounting.contacts'
  ],
});

let tenantId = null;

// Root route
app.get('/', (req, res) => {
  res.send(`
    <h2>Welcome to the Xero Integration App</h2>
    <p><a href="/connect">Connect to Xero</a></p>
  `);
});

// Add the /connect route that starts the OAuth flow
app.get('/connect', async (req, res) => {
  try {
    const consentUrl = await xero.buildConsentUrl();
    res.redirect(consentUrl);
  } catch (error) {
    console.error('Error building consent URL:', error);
    res.status(500).send('Failed to start OAuth flow');
  }
});

app.listen(port, () => {
  console.log(`Server running at http://localhost:${port}`);
});
