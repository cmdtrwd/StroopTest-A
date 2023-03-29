// Import I/O module
const fs = require('fs');

// Set the base file name and extension
const baseFileName = "result";
const extension = ".xlsx";

// Set the initial file number
let fileNumber = 1;

// Loop until a unique file name is found
while (fs.existsSync(`output/${baseFileName}_${fileNumber}${extension}`)) {
    fileNumber++;
}

//construct the file name
const fileName = `output/${baseFileName}_${fileNumber}${extension}`;

// Import xlsx module
const XLSX = require("xlsx");
const jsontoxml = require("jsontoxml");

// Create a new workbook
const workbook = XLSX.utils.book_new();

// Insert data into a worksheet
function InsertData() {
    
    const worksheet = XLSX.utils.json_to_sheet([
        {
            "First Name": "John",
            "Last Name": "Jack",
            "Gender": "Male",
            "Country": "United States",
            "Age": 28,
            "Date": "21/09/2022",
            "ID": 16001
        },
        {
            "First Name": "Sarah",
            "Last Name": "Fin",
            "Gender": "Female",
            "Country": "United States",
            "Age": 30,
            "Date": "22/09/2022",
            "ID": 16002
        },
        {
            "First Name": "Bob",
            "Last Name": "Maley",
            "Gender": "Male",
            "Country": "Thailand",
            "Age": 18,
            "Date": "23/09/2022",
            "ID": 16003
        },
    ]);

    XLSX.utils.book_append_sheet(workbook, worksheet, "Data");
    XLSX.writeFile(workbook, fileName);
}

InsertData();

// Import path module
const path = require('path');

// Import express module (After installing express module using npm)
const express = require('express');
const app = express();

// Serve static files from the "public" directory
app.use(express.static('public'));

// using path module to get a response from a defined path

const savePage = path.join(__dirname, '../templates/filesaved.html');

app.get('/home', function (req, res) {
    res.status(200);
    res.type('text/html');
    res.sendFile(savePage);
});

app.listen(3000, function(){
    console.log("Server is listening on port 3000");
})
