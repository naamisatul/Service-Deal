const express = require('express');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');
const cors = require('cors');
const bodyParser = require('body-parser');

const app = express();
const PORT = 3000;

// Middleware
app.use(cors());
app.use(bodyParser.json());
app.use(express.static(path.join(__dirname)));

// Excel file path
const EXCEL_FILE = path.join(__dirname, 'bookings.xlsx');

// Function to initialize Excel file if it doesn't exist
function initializeExcel() {
    if (!fs.existsSync(EXCEL_FILE)) {
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.aoa_to_sheet([['Name', 'Phone', 'Service', 'Timestamp']]);
        XLSX.utils.book_append_sheet(wb, ws, 'Bookings');
        XLSX.writeFile(wb, EXCEL_FILE);
        console.log('Excel file initialized');
    }
}

// Initialize Excel file on startup
initializeExcel();

// Route to handle booking submissions
app.post('/api/book', (req, res) => {
    try {
        const { name, phone, service } = req.body;

        console.log('Received booking request:', { name, phone, service });

        if (!name || !phone || !service) {
            console.log('Missing required fields');
            return res.status(400).json({ error: 'Missing required fields: name, phone, or service' });
        }

        // Validate phone number format
        if (!/^[0-9]{10}$/.test(phone)) {
            console.log('Invalid phone number format:', phone);
            return res.status(400).json({ error: 'Invalid phone number format. Please enter a 10-digit number.' });
        }

        // Read existing Excel file
        if (!fs.existsSync(EXCEL_FILE)) {
            console.log('Excel file not found, creating new one');
            initializeExcel();
        }

        const wb = XLSX.readFile(EXCEL_FILE);
        const ws = wb.Sheets['Bookings'];

        // Get existing data
        const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

        // Ensure we have headers
        if (data.length === 0 || !data[0].includes('Name')) {
            data.unshift(['Name', 'Phone', 'Service', 'Timestamp']);
        }

        // Add new row
        const timestamp = new Date().toISOString();
        data.push([name, phone, service, timestamp]);

        // Create new worksheet
        const newWs = XLSX.utils.aoa_to_sheet(data);
        wb.Sheets['Bookings'] = newWs;

        // Write back to file
        XLSX.writeFile(wb, EXCEL_FILE);

        console.log(`✅ Booking saved successfully: ${name}, ${phone}, ${service}`);
        res.json({
            success: true,
            message: 'Booking saved successfully',
            data: { name, phone, service, timestamp }
        });

    } catch (error) {
        console.error('❌ Error saving booking:', error);
        res.status(500).json({
            error: 'Failed to save booking to Excel file',
            details: error.message
        });
    }
});

// Start server
app.listen(PORT, () => {
    console.log(`Server running on http://localhost:${PORT}`);
    console.log('Excel file will be created/updated at:', EXCEL_FILE);
});