const express = require('express');
const axios = require('axios');
const multer = require('multer');
const path = require('path');
const bodyParser = require('body-parser');
const fs = require('fs');
const { DateTime } = require('luxon');

const app = express();
const upload = multer({ dest: 'uploads/' });

// Xero app credentials
const CLIENT_ID = '93B4125E75714BCCBB2B3CAB4E5AC7CF';
const CLIENT_SECRET = 'm_2kdSW-3SkAzHYaIvPt8yhzSEkVk8myxfhIVrnAgzkBD9Za';
const REDIRECT_URI = 'http://localhost:5000/callback';
const SCOPES = 'openid profile email accounting.transactions offline_access';

let tokens = {};
let tenant_id = null;

// Middleware for JSON handling
app.use(bodyParser.json());
app.set('view engine', 'ejs');

// Step 1: Redirect to Xero for Authorization
app.get('/', (req, res) => {
    const authUrl = `https://login.xero.com/identity/connect/authorize?` +
                    `response_type=code&` +
                    `client_id=${CLIENT_ID}&` +
                    `redirect_uri=${encodeURIComponent(REDIRECT_URI)}&` +
                    `scope=${encodeURIComponent(SCOPES)}&` +
                    `state=12345`;
    res.redirect(authUrl);
});

// Step 2: Callback to get authorization code and exchange tokens
app.get('/callback', async (req, res) => {
    const error = req.query.error;
    if (error) {
        return res.send(`❌ Authorization failed: ${error}`);
    }

    const code = req.query.code;
    if (!code) {
        return res.send('⚠️ No authorization code received.');
    }

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

        // Get tenant ID
        const connectionsResponse = await axios.get('https://api.xero.com/connections', {
            headers: {
                'Authorization': `Bearer ${tokens.access_token}`
            }
        });

        const tenants = connectionsResponse.data;
        if (tenants.length === 0) {
            return res.send('❌ No Xero tenant found.');
        }

        tenant_id = tenants[0].tenantId;
        res.send(`
            ✅ Authorization successful!<br><br>
            Now upload your Excel file at <a href="/upload">/upload</a>
        `);
    } catch (error) {
        console.error(error);
        res.send(`❌ Failed to get access token: ${error.message}`);
    }
});

// Upload Excel file and build PO
app.get('/upload', (req, res) => {
    res.render('upload');  // Render a file upload form (we'll create this template)
});

app.post('/upload', upload.single('file'), (req, res) => {
    const file = req.file;
    if (!file) {
        return res.status(400).send('No file uploaded');
    }

    // Parse Excel file here, using a library like 'xlsx' or 'exceljs'
    // Here's a simplified approach using 'xlsx' for demonstration
    const xlsx = require('xlsx');
    const workbook = xlsx.readFile(file.path);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });

    // Extract information and build PO (similar to Python code logic)
    const lineItems = extractLineItems(data);
    const quoteInfo = extractQuoteInfo(data);

    // Prepare PO data
    const poData = buildPoData(lineItems, quoteInfo);

    // Store PO data for later sending to Xero
    global.poPayload = poData;

    res.render('po', { poJson: JSON.stringify(poData, null, 4) });
});

function extractLineItems(data) {
    const descriptions = extractColumnValues(data, "Description") || [];
    const quantities = extractColumnValues(data, "Qty") || [];
    const unitPrices = extractColumnValues(data, "Unit Price") || [];

    return descriptions.map((desc, i) => ({
        Description: desc,
        Quantity: parseFloat(quantities[i] || 1),
        UnitAmount: parseFloat(unitPrices[i] || 0),
        AccountCode: '400',
        TaxType: 'INPUT'
    }));
}

function extractColumnValues(data, fieldName) {
    for (let rowIdx = 0; rowIdx < data.length; rowIdx++) {
        for (let colIdx = 0; colIdx < data[rowIdx].length; colIdx++) {
            if (String(data[rowIdx][colIdx]).toLowerCase() === fieldName.toLowerCase()) {
                return data.slice(rowIdx + 1).map(row => row[colIdx]).filter(value => value);
            }
        }
    }
    return null;
}

function extractQuoteInfo(data) {
    let quoteInfo = {};
    const sectionTitle = 'QUOTE INFORMATION';
    for (let i = 0; i < data.length; i++) {
        if (data[i].includes(sectionTitle)) {
            const startRow = i + 1;
            for (let j = startRow; j < startRow + 10; j++) {
                if (data[j]) {
                    for (let k = 0; k < data[j].length; k += 2) {
                        const key = String(data[j][k]).trim();
                        const value = String(data[j][k + 1]).trim();
                        if (key) {
                            quoteInfo[key] = value;
                        }
                    }
                }
            }
            break;
        }
    }
    return quoteInfo;
}

function buildPoData(lineItems, quoteInfo) {
    const contactName = quoteInfo["Reseller Contact"] || "Unknown Supplier";
    const reference = quoteInfo["Sales Quotation"] || "AutoPO";
    const currencyCode = quoteInfo["Currency"] || "AUD";
    const rawDate = quoteInfo["Validity End Date"];
    const deliveryDate = rawDate ? DateTime.fromFormat(rawDate, 'dd/MM/yyyy').toISODate() : DateTime.now().toISODate();

    return {
        Contact: { Name: contactName },
        Date: DateTime.now().toISODate(),
        DeliveryDate: deliveryDate,
        LineItems: lineItems,
        Reference: reference,
        CurrencyCode: currencyCode,
        Status: 'DRAFT'
    };
}

// Send the PO to Xero
app.post('/send_po', async (req, res) => {
    if (!global.poPayload) {
        return res.status(400).send('❌ No PO payload available. Upload an Excel file first.');
    }

    try {
        const response = await axios.post('https://api.xero.com/api.xro/2.0/PurchaseOrders', global.poPayload, {
            headers: {
                'Authorization': `Bearer ${tokens.access_token}`,
                'Xero-tenant-id': tenant_id,
                'Content-Type': 'application/json'
            }
        });

        res.send(`✅ Purchase Order sent successfully!<br><pre>${JSON.stringify(response.data, null, 2)}</pre>`);
    } catch (error) {
        res.status(500).send(`❌ Failed to send PO: ${error.message}`);
    }
});

// Start the server
app.listen(5000, () => {
    console.log('Server running on http://localhost:5000');
});
