const express = require("express");
const multer = require("multer");
const xlsx = require("xlsx");
const axios = require("axios");
const fs = require("fs");
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
      const data = xlsx.utils.sheet_to_json(worksheet, {
        header: 1,
        defval: "",
      });
      console.log(`Processing sheet: ${sheetName}`);

      const headers = data[0];
      const vehicleDetailsIndex = headers.indexOf("VEHICLE DETAILS");

      if (vehicleDetailsIndex === -1) {
        return res.status(400).send("VEHICLE DETAILS column not found");
      }

      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        let vehicleDetails = row[vehicleDetailsIndex];

        if (vehicleDetails == null || vehicleDetails.trim() === "") {
          // row[vehicleDetailsIndex] = "Invalid or missing VEHICLE DETAILS";
          console.log(`Row ${i + 1}: Invalid or missing VEHICLE DETAILS`);
        } else if (
          typeof vehicleDetails === "string" &&
          isVIN(vehicleDetails.trim())
        ) {
          console.log(`Row ${i + 1}: Valid VIN found: ${vehicleDetails}`);
          try {
            const response = await axios.get(
              `https://vpic.nhtsa.dot.gov/api/vehicles/DecodeVinValuesExtended/${vehicleDetails.trim()}?format=json`
            );

            if (response.data.Results && response.data.Results.length > 0) {
              const carDetails = response.data.Results[0];

              // Replace VIN with car make, model, and year
              const make = carDetails.Make || "Unknown Make";
              const model = carDetails.Model || "Unknown Model";
              const year = carDetails.ModelYear || "Unknown Year";
              row[vehicleDetailsIndex] = `${make} ${model} (${year})`;
              console.log(
                `Row ${i + 1}: Replaced VIN with: ${make} ${model} (${year})`
              );
            } else {
              row[vehicleDetailsIndex] = "Unknown Car";
              console.log(
                `Row ${i + 1}: No records found, marked as Unknown Car`
              );
            }
          } catch (error) {
            console.error(
              `Row ${i + 1}: Error fetching data for VIN: ${vehicleDetails}`,
              error
            );
            // row[vehicleDetailsIndex] = `Error fetching data for VIN: ${vehicleDetails}`;
          }
        } else if (typeof vehicleDetails === "string") {
          console.log(
            `Row ${i + 1}: Non-VIN entry, leaving as is: ${vehicleDetails}`
          );
          // Keep the entry as it is for non-VIN strings
        } else {
          // Set an appropriate message for invalid or missing VEHICLE DETAILS
          // row[vehicleDetailsIndex] = "Invalid or missing VEHICLE DETAILS";
          console.log(`Row ${i + 1}: Invalid or missing VEHICLE DETAILS`);
        }
      }

      // Update the worksheet with the new data
      const newWorksheet = xlsx.utils.aoa_to_sheet(data);
      workbook.Sheets[sheetName] = newWorksheet;
    }

    // Write the updated workbook back to the same file
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
      // Do not delete the file after download
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
  return /^[A-HJ-NPR-Z0-9]{17}$/.test(vin.trim()); // Simple VIN validation
}

// Start the server
app.listen(port, () => {
  console.log(`Server running on port ${port}`);
});
