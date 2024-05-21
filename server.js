const express = require('express');
const bodyParser = require('body-parser');
const fs = require('fs');
const XLSX = require('xlsx');

const app = express();
const port = 3000;

app.use(bodyParser.json());

app.post('/store-data', (req, res) => {
    const { location, bloodType } = req.body;

    // Load existing data
    let workbook;
    let sheet;
    try {
        workbook = XLSX.readFile('donors.xlsx');
        sheet = workbook.Sheets['Sheet1'];
    } catch (error) {
        workbook = XLSX.utils.book_new();
        sheet = XLSX.utils.json_to_sheet([]);
        XLSX.utils.book_append_sheet(workbook, sheet, 'Sheet1');
    }

    // Append new data
    const newRow = { Location: location, BloodType: bloodType };
    const data = XLSX.utils.sheet_to_json(sheet);
    data.push(newRow);
    const newSheet = XLSX.utils.json_to_sheet(data);

    // Update workbook
    workbook.Sheets['Sheet1'] = newSheet;
    XLSX.writeFile(workbook, 'donors.xlsx');

    res.sendStatus(200);
});

app.listen(port, () => {
    console.log(`Server running at http://localhost:${port}`);
});
