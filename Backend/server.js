const express = require("express");
const multer = require("multer");
const xlsx = require("xlsx");
const axios = require("axios");
const app = express();
const port = 3003;

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
            const response = await axios.get(`https://vpic.nhtsa.dot.gov/api/vehicles/DecodeVinValuesExtended/${vehicleDetails}?format=json`);

            if (response.data.Results && response.data.Results.length > 0) {
              const carDetails = response.data.Results[0];

              // Replace VIN with car make, model, and year
              const make = carDetails.Make;
              const model = carDetails.Model;
              const year = carDetails.ModelYear;
              row[vehicleDetailsIndex] = `${make} ${model} (${year})`;
              console.log(`Replaced VIN with: ${make} ${model} (${year})`);
            } else {
              row[vehicleDetailsIndex] = "Unknown Car";
              console.log("No records found, marked as Unknown Car");
            }
          } catch (error) {
            console.error(`Error fetching data for VIN: ${vehicleDetails}`, error);
            row[vehicleDetailsIndex] = `Error fetching data for VIN: ${vehicleDetails}`;
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

