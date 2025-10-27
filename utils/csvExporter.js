
const fs = require('fs');
const path = require('path');
const { Parser } = require('json2csv');

function generateTimestamp() {
    const now = new Date();
    const date = now.toISOString().split('T')[0];
    const time = now.toTimeString().split(' ')[0].replace(/:/g, '-');
    return `${date}_${time}`;
}

// Dynamically infer fields from data
function inferFieldsFromData(data) {
    const fieldSet = new Set();
    data.forEach(item => {
        Object.keys(item).forEach(key => fieldSet.add(key));
    });
    return [...fieldSet];
}

// Exports CSV and saves it to disk
async function exportToCSVFile(data, baseName) {
    if (!data || !Array.isArray(data) || data.length === 0) {
        throw new Error('No data provided for CSV export.');
    }

    const fields = inferFieldsFromData(data);
    const parser = new Parser({ fields });
    const csv = parser.parse(data);

    const timestamp = generateTimestamp();
    const fileName = `${baseName}_${timestamp}.csv`;
    const filePath = path.join(process.cwd(), '..', 'downloads', fileName);

    // Auto-create the downloads directory if it doesn’t exist
    if (!fs.existsSync(path.dirname(filePath))) {
        fs.mkdirSync(path.dirname(filePath), { recursive: true });
    }

    fs.writeFileSync(filePath, csv);
    return filePath;
}

// Exports CSV as string (for email attachment)
async function exportToCSVString(data, baseName) {
    if (!data || !Array.isArray(data) || data.length === 0) {
        throw new Error('No data provided for CSV export.');
    }

    const fields = inferFieldsFromData(data);
    const parser = new Parser({ fields });
    const csv = parser.parse(data);

    const timestamp = generateTimestamp();
    const filename = `${baseName}_${timestamp}.csv`;

    return { csv, filename };
}

module.exports = {
    exportToCSVFile,
    exportToCSVString
};

