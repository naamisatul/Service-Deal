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
    const { name, phone, service } = req.body;

    if (!name || !phone || !service) {
        return res.status(400).json({ error: 'Missing required fields' });
    }

    try {
        // Read existing Excel file
        const wb = XLSX.readFile(EXCEL_FILE);
        const ws = wb.Sheets['Bookings'];

        // Get existing data
        const data = XLSX.utils.sheet_to_json(ws, { header: 1 });

        // Add new row
        const timestamp = new Date().toISOString();
        data.push([name, phone, service, timestamp]);

        // Create new worksheet
        const newWs = XLSX.utils.aoa_to_sheet(data);
        wb.Sheets['Bookings'] = newWs;

        // Write back to file
        XLSX.writeFile(wb, EXCEL_FILE);

        console.log(`New booking added: ${name}, ${phone}, ${service}`);
        res.json({ success: true, message: 'Booking saved successfully' });

    } catch (error) {
        console.error('Error saving booking:', error);
        res.status(500).json({ error: 'Failed to save booking' });
    }
});

// Start server
app.listen(PORT, () => {
    console.log(`Server running on http://localhost:${PORT}`);
    console.log('Excel file will be created/updated at:', EXCEL_FILE);
});