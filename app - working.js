const express = require('express');
const axios = require('axios');
const multer = require('multer');
const path = require('path');
const bodyParser = require('body-parser');
const fs = require('fs');
const { DateTime } = require('luxon');
const xlsx = require('xlsx');

const app = express();
const upload = multer({ dest: 'uploads/' });

// Xero app credentials
const CLIENT_ID = '93B4125E75714BCCBB2B3CAB4E5AC7CF';
const CLIENT_SECRET = 'm_2kdSW-3SkAzHYaIvPt8yhzSEkVk8myxfhIVrnAgzkBD9Za';
const REDIRECT_URI = 'https://xero-integration-p55k.onrender.com/callback';
const SCOPES = 'openid profile email accounting.transactions offline_access';

let tokens = {};
let tenant_id = null;
let poPayload = null;

// Refresh access token using the refresh_token
async function refreshAccessToken() {
    if (!tokens.refresh_token) {
        throw new Error("No refresh token available");
    }

    try {
        const response = await axios.post('https://identity.xero.com/connect/token',
            new URLSearchParams({
                grant_type: 'refresh_token',
                refresh_token: tokens.refresh_token,
                client_id: CLIENT_ID,
                client_secret: CLIENT_SECRET
            }).toString(),
            {
                headers: {
                    'Content-Type': 'application/x-www-form-urlencoded'
                }
            }
        );

        tokens = response.data;
        console.log("🔄 Access token refreshed");
    } catch (error) {
        console.error("❌ Failed to refresh token:", error.response?.data || error.message);
        throw new Error("Failed to refresh token");
    }
}

// Middleware
app.use(bodyParser.json());
app.set('view engine', 'ejs');

// Step 1: Auth redirect
app.get('/', (req, res) => {
    const authUrl = `https://login.xero.com/identity/connect/authorize?` +
        `response_type=code&` +
        `client_id=${CLIENT_ID}&` +
        `redirect_uri=${encodeURIComponent(REDIRECT_URI)}&` +
        `scope=${encodeURIComponent(SCOPES)}&` +
        `state=12345`;
    res.redirect(authUrl);
});

// Step 2: Handle callback
app.get('/callback', async (req, res) => {
    const error = req.query.error;
    if (error) return res.send(`❌ Authorization failed: ${error}`);

    const code = req.query.code;
    if (!code) return res.send('⚠️ No authorization code received.');

    try {
        const tokenResponse = await axios.post('https://identity.xero.com/connect/token', new URLSearchParams({
            grant_type: 'authorization_code',
            code,
            redirect_uri: REDIRECT_URI,
            client_id: CLIENT_ID,
            client_secret: CLIENT_SECRET
        }).toString(), {
            headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
        });

        tokens = tokenResponse.data;

        const connectionsResponse = await axios.get('https://api.xero.com/connections', {
            headers: {
                'Authorization': `Bearer ${tokens.access_token}`
            }
        });

        const tenants = connectionsResponse.data;
        if (tenants.length === 0) return res.send('❌ No Xero tenant found.');

        tenant_id = tenants[0].tenantId;
        res.send(`✅ Authorization successful!<br><br>Now upload your Excel file at <a href="/upload">/upload</a>`);
    } catch (error) {
        console.error(error);
        res.send(`❌ Failed to get access token: ${error.message}`);
    }
});

// Upload form
app.get('/upload', (req, res) => {
    res.render('upload');
});

// Handle Excel upload
app.post('/upload', upload.single('file'), async (req, res) => {
    const file = req.file;
    if (!file) return res.status(400).send('No file uploaded');

    const workbook = xlsx.readFile(file.path);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });

    const extractColumnValues = (fieldName) => {
        for (let rowIdx = 0; rowIdx < data.length; rowIdx++) {
            for (let colIdx = 0; colIdx < data[rowIdx].length; colIdx++) {
                const cellValue = String(data[rowIdx][colIdx]).trim();
                if (cellValue.toLowerCase() === fieldName.toLowerCase()) {
                    const values = [];
		    let currentRow = rowIdx + 1;
                    while (currentRow < data.length) {
			const value = data[currentRow][colIdx];
                    	if (value === null || value === undefined || value === '') {
                           break;
			}
                        values.push(value);
                        currentRow++;
                    	}
                    	return values;
                   }
            	}
	    }
        return null;
    };

    const extractQuoteInfo = (sectionTitle = "QUOTE INFORMATION") => {
        const quoteInfo = {};
        for (let i = 0; i < data.length; i++) {
            for (let j = 0; j < data[i].length; j++) {
                const cell = String(data[i][j]).trim();
                if (cell === sectionTitle) {
                    for (let k = i + 1; k < i + 10 && k < data.length; k++) {
                        const row = data[k].filter(val => val != null && val !== '');
                        for (let idx = 0; idx < row.length - 1; idx += 2) {
                            quoteInfo[row[idx]] = row[idx + 1];
                        }
                    }
                    return quoteInfo;
                }
            }
        }
        return {};
    };

    const itemNumbers = extractColumnValues("Item Number") || [];
    const descriptions = extractColumnValues("Description") || [];
    const quantities = extractColumnValues("Qty") || [];
    const unitPrices = extractColumnValues("Unit Price") || [];
    const quoteInfo = extractQuoteInfo();

    const contactName = quoteInfo["Reseller Contact"] || "Unknown Supplier";
    const reference = quoteInfo["Sales Quotation"] || "AutoPO";
    let currencyCode = quoteInfo["Currency"] || "AUD";
    if (!["AUD", "NZD"].includes(currencyCode)) currencyCode = "AUD";

    const rawDate = quoteInfo["Validity End Date"] || "";
    let deliveryDate = DateTime.now().toISODate();
    try {
        deliveryDate = rawDate ? DateTime.fromFormat(rawDate, "dd/MM/yyyy").toISODate() : deliveryDate;
    } catch {}

    const lineItems = itemNumbers.map((_, i) => ({
        Description: String(descriptions[i]) || '',
        Quantity: parseFloat(quantities[i]) || 0,
        UnitAmount: parseFloat(unitPrices[i]) || 0,
        AccountCode: '400',
        TaxType: 'INPUT'
    }));

    // Get or create ContactID
    async function get_or_create_contact_id(contactName, retried = false) {
        try {
            const response = await axios.get('https://api.xero.com/api.xro/2.0/Contacts', {
                headers: {
                    'Authorization': `Bearer ${tokens.access_token}`,
                    'Xero-tenant-id': tenant_id
                },
                params: { name: contactName }
            });

            if (response.data.Contacts.length > 0) {
                return response.data.Contacts[0].ContactID;
            } else {
                const newContactResponse = await axios.post(
                    'https://api.xero.com/api.xro/2.0/Contacts',
                    {
                        Contacts: [
                            { Name: contactName }
                        ]
                    },
                    {
                        headers: {
                            'Authorization': `Bearer ${tokens.access_token}`,
                            'Xero-tenant-id': tenant_id,
                            'Content-Type': 'application/json'
                        }
                    }
                );

                return newContactResponse.data.Contacts[0].ContactID;
            }
        } catch (error) {
            if (error.response?.status === 401 && !retried) {
                console.log("🔐 Token expired. Attempting to refresh...");
                await refreshAccessToken();
                return await get_or_create_contact_id(contactName, true);
            }
            console.error("❌ Error in get_or_create_contact_id:", error.response?.data || error.message);
            throw new Error(error.response?.data?.Message || "Error getting/creating contact");
        }
    }

    let contactId;
    try {
        contactId = await get_or_create_contact_id(contactName);
    } catch (e) {
        return res.status(400).send(`❌ Error getting/creating contact: ${e.message}`);
    }

    const toXeroDateFormat = (date) => `/Date(${DateTime.fromISO(date).toMillis()}+0000)/`;

    poPayload = {
        Contact: {
            ContactID: contactId,
            Name: contactName
        },
        Date: toXeroDateFormat(DateTime.now().toISO()),
        DeliveryDate: toXeroDateFormat(deliveryDate),
        LineItems: lineItems,
        DeliveryAddress: "Enablis Office",
        Reference: reference,
        CurrencyCode: currencyCode,
        Status: "DRAFT"
    };

    res.render('po', { poJson: JSON.stringify(poPayload, null, 4) });
});

// Send PO to Xero
app.post('/send_po', async (req, res) => {
    if (!poPayload) return res.status(400).send('❌ No PO payload available. Upload an Excel file first.');

    try {
        const response = await axios.post('https://api.xero.com/api.xro/2.0/PurchaseOrders', poPayload, {
            headers: {
                'Authorization': `Bearer ${tokens.access_token}`,
                'Xero-tenant-id': tenant_id,
                'Content-Type': 'application/json'
            }
        });

        res.send(`✅ Purchase Order sent successfully!<br><pre>${JSON.stringify(response.data, null, 2)}</pre>`);
    } catch (error) {
        if (error.response) {
            const apiError = error.response.data;
            console.error('❌ Xero API Error Response:', JSON.stringify(apiError, null, 2));
            const validationErrors = apiError.Elements?.[0]?.ValidationErrors || [];
            const errorMessages = validationErrors.map(err => `- ${err.Message}`).join('\n');

            return res.status(400).send(`
                ❌ Xero Validation Error(s):<br><br>
                <pre>${errorMessages || apiError.Message}</pre>
            `);
        } else {
            console.error('❌ Request Error:', error.message);
            res.status(500).send(`❌ Failed to send PO: ${error.message}`);
        }
    }
});

// Start the server
app.listen(5000, () => {
    console.log('🚀 Server running on http://localhost:5000');
});
