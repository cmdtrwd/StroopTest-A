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

// // Extract encoded data from URL query parameter
// const queryString = window.location.search;
// const urlParams = new URLSearchParams(queryString);
// const encodedData = urlParams.get("data");

// // Parse encoded data into JSON object
// const data = JSON.parse(decodeURIComponent(encodedData));




// Import path module
const path = require('path');

// Import express module (After installing express module using npm)
const express = require('express');
const { url } = require('inspector');
const { endianness } = require('os');
const app = express();

// Serve static files from the "public" directory
app.use(express.static('public'));
app.use(express.json());


// Import cors module
const cors = require('cors');
const corsOptions = {
    origin: 'http://localhost:3000/',
    // origin: null,
    credentials: true,
};
app.use(cors(corsOptions));


// Insert data into a worksheet
    
app.post('/send-data', (req, res)=>{
    // const data = JSON.stringify(req.body);
    // const worksheet = XLSX.utils.json_to_sheet(data);
    // XLSX.utils.book_append_sheet(workbook, worksheet, "Data");
    // XLSX.writeFile(workbook, fileName);
    res.send("Data received");    
});




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
