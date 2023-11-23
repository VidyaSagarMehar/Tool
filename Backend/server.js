// server.js

const express = require('express');
const bodyParser = require('body-parser');
const exceljs = require('exceljs');
const fs = require('fs');
const cors = require('cors');

const app = express();
const port = process.env.PORT || 5000;

app.use(bodyParser.json());
app.use(cors());

app.get('/', (req, res) => {
	res.sendFile(__dirname + '/index.html'); // Adjust the path as needed
});

app.post('/storeData', (req, res) => {
	const data = req.body;
	const filePath = 'path/to/store/data.xlsx';

	// Create a new workbook or open the existing one
	const workbook = new exceljs.Workbook();

	// Check if the file exists
	const fileExists = fs.existsSync(filePath);

	workbook.xlsx
		.readFile(filePath)
		.then(() => {
			const worksheet = workbook.getWorksheet('Data');

			// If the file didn't exist, add a new worksheet
			if (!fileExists || !worksheet) {
				workbook.addWorksheet('Data');
			}
			// Retrieve the worksheet again
			const updatedWorksheet = workbook.getWorksheet('Data');

			// const worksheet = workbook.getWorksheet('Data');

			// Add headers if the worksheet is empty
			if (updatedWorksheet.rowCount) {
				updatedWorksheet.columns = [
					{ header: 'Date', key: 'date' },
					{ header: 'Platform', key: 'platform' },
					{ header: 'Email', key: 'email' },
					{ header: 'Subject', key: 'subject' },
					{ header: 'Mail_type', key: 'mailType' },
					{ header: 'Sme_name', key: 'smeName' },
					{ header: 'Mai;_status', key: 'mailStatus' },
				];
			}

			// Add a new row with the data
			updatedWorksheet.addRow(data);

			// Save the workbook
			return workbook.xlsx.writeFile(filePath);
		})
		.then(() => {
			console.log('Data saved to Excel:', filePath);
			res.json({ message: 'Data saved to Excel' });
		})
		.catch((error) => {
			console.error('Error saving data to Excel:', error);
			res.status(500).json({ message: 'Error saving data to Excel' });
		});
});

// app.get('/fetchData', (req, res) => {
// 	// Read data from the Excel file and send it as JSON
// 	// Include appropriate error handling
// 	// Adjust the file path as needed
// 	const filePath = 'path/to/store/data.xlsx';
// 	const workbook = new exceljs.Workbook();

// 	workbook.xlsx
// 		.readFile(filePath)
// 		.then(() => {
// 			const worksheet = workbook.getWorksheet('Data');
// 			const data = [];

// 			worksheet.eachRow({ includeEmpty: false }, (row) => {
// 				data.push({
// 					email: row.values[1],
// 					date: row.values[2],
// 					// Add other fields as needed
// 				});
// 			});

// 			res.json(data);
// 		})
// 		.catch((error) => {
// 			console.error('Error reading data from Excel:', error);
// 			res.status(500).json({ message: 'Error reading data from Excel' });
// 		});
// });

app.get('/getAllData', (req, res) => {
	const filePath = 'path/to/store/data.xlsx';
	const workbook = new exceljs.Workbook();

	workbook.xlsx
		.readFile(filePath)
		.then(() => {
			const worksheet = workbook.getWorksheet('Data');

			// Convert worksheet data to JSON
			const jsonData = [];
			const headerRow = worksheet.getRow(1);

			worksheet.eachRow((row, rowNumber) => {
				if (rowNumber !== 1) {
					// Skip the header row
					const rowData = {};
					row.eachCell((cell, colNumber) => {
						const columnHeader = headerRow.getCell(colNumber).value;
						rowData[columnHeader] = cell.value;
					});
					jsonData.push(rowData);
				}
			});

			res.json(jsonData);
		})
		.catch((error) => {
			console.error('Error getting data from Excel:', error);
			res.status(500).json({ message: 'Error getting data from Excel' });
		});
});

app.listen(port, () => {
	console.log(`Server is running on port ${port}`);
});
