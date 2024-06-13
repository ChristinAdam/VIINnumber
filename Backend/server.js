const express = require("express");
const multer = require("multer");
const xlsx = require("xlsx");
const axios = require("axios");
const fs = require("fs");
const path = require("path");
const app = express();
const port = 3003;

// VIN API configuration
const VIN_Url = "https://auto.dev/api/listings";
const API_key = "ZrQEPSkKY2hyaXN0aW5hZGFtejc4NkBnbWFpbC5jb20=";

// Multer storage configuration
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, "uploads/");
  },
  filename: (req, file, cb) => {
    cb(null, file.originalname);
  },
});

const upload = multer({ storage: storage });

app.post("/upload", upload.single("file"), async (req, res) => {
  try {
    console.log("File received:", req.file.path);
    const workbook = xlsx.readFile(req.file.path);

    // Iterate over all sheets in the workbook
    for (const sheetName of workbook.SheetNames) {
      const worksheet = workbook.Sheets[sheetName];
      const data = xlsx.utils.sheet_to_json(worksheet, { header: 1 });
      console.log(`Processing sheet: ${sheetName}`);

      const headers = data[0];
      const vehicleDetailsIndex = headers.indexOf("VEHICLE DETAILS");

      if (vehicleDetailsIndex === -1) {
        return res.status(400).send("VEHICLE DETAILS column not found");
      }

      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const vehicleDetails = row[vehicleDetailsIndex];

        if (typeof vehicleDetails === "string" && isVIN(vehicleDetails)) {
          console.log(`Valid VIN found: ${vehicleDetails}`);
          try {
            const response = await axios.get(`${VIN_Url}?vin=${vehicleDetails}&apikey=${API_key}`);
            const records = response.data.records;

            if (records && records.length > 0) {
              const carDetails = records[0];

              // Replace VIN with car make and model
              if (carDetails.make && carDetails.model) {
                row[vehicleDetailsIndex] = `${carDetails.make} ${carDetails.model}`;
                console.log(`Replaced VIN with: ${carDetails.make} ${carDetails.model}`);
              } else {
                row[vehicleDetailsIndex] = "Unknown Car";
                console.log("Car details not found, marked as Unknown Car");
              }
            } else {
              row[vehicleDetailsIndex] = "Unknown Car";
              console.log("No records found, marked as Unknown Car");
            }
          } catch (error) {
            console.error(`Error fetching data for VIN: ${vehicleDetails}`, error);
            row[vehicleDetailsIndex] = "Error fetching car details";
          }
        } else {
          console.log(`Invalid or missing VIN: ${vehicleDetails}`);
        }
      }

      // Create new worksheet with updated data
      const newWorksheet = xlsx.utils.aoa_to_sheet(data);
      workbook.Sheets[sheetName] = newWorksheet;
    }

    // Write the entire workbook to file once all sheets are processed
    xlsx.writeFile(workbook, req.file.path);
    console.log("Processing completed");

    // Send the updated file to the client
    res.download(req.file.path, (err) => {
      if (err) {
        console.error("Error downloading updated file", err);
        if (!res.headersSent) {
          res.status(500).send("Error downloading updated file");
        }
      }
    });
  } catch (error) {
    console.error("Error processing file", error);
    if (!res.headersSent) {
      res.status(500).send("Error processing file");
    }
  }
});

// Function to validate VIN
function isVIN(vin) {
  return /^[A-HJ-NPR-Z0-9]{17}$/.test(vin); // Simple VIN validation
}

// Start the server
app.listen(port, () => {
  console.log(`Server running on port ${port}`);
});
