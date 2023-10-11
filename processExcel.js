const xlsx = require("xlsx");
const fs = require("fs");
const path = require("path");

const ENDPOINT =
  "https://generativeai.aiservice.us-chicago-1.oci.oraclecloud.com";
const COMPARTMENT_ID =
  "ocid1.compartment.oc1..aaaaaaaaxdgp4qx2uodttcglrsg7b24vkuydba7hbsh6l3meaovxfw4ceybq";
const MODEL_ID = "cohere.command";
const MAX_TOKENS = 100;
const STOP_SEQUENCE = "Input:";
const TEMPERATURE = 0;
const MAX_RPM = 15; //Maximum requests per minute to fire to GenAI (Cohere) service

// Check if file path and cell range are provided
if (process.argv.length <= 3) {
  console.log(
    "Please provide the Excel file path and cell range as arguments."
  );
  process.exit(1);
}

const filePath = process.argv[2];
const cellRangeToProcess = process.argv[3];

// Validate cell range format
if (!/^[A-Za-z ]+![A-Z]+\d+:[A-Z]+\d+$/.test(cellRangeToProcess)) {
  console.log(
    "Invalid cell range format. Please use the format 'Worksheet Name!Cell:Cell'."
  );
  process.exit(1);
}

function generateText(prompt) {
  try {
    console.log("/generateText start");
    const servingMode = {
      modelId: MODEL_ID,
      servingType: "ON_DEMAND",
    };
    const generateTextDetails = {
      prompts: [prompt],
      servingMode: servingMode,
      compartmentId: COMPARTMENT_ID,
      maxTokens: MAX_TOKENS,
      temperature: TEMPERATURE,
      //frequencyPenalty: 1.0,
      //topP: 0.7,
    };
    if (STOP_SEQUENCE !== undefined && STOP_SEQUENCE !== null) {
      console.log("adding stopSequences", STOP_SEQUENCE);
      generateTextDetails.stopSequences = [STOP_SEQUENCE];
    }
    console.log("message for Cohere generation", generateTextDetails);
    const generateTextRequest = {
      generateTextDetails: generateTextDetails,
    };
    /*
        const CONFIG_PROFILE = "DEFAULT";
        provider = new common.ConfigFileAuthenticationDetailsProvider("config", CONFIG_PROFILE);
        const generateAiClient = new generative_ai.GenerativeAiClient({
            authenticationDetailsProvider: provider,
        });
        generateAiClient.endpoint = ENDPOINT;
        const generateTextResponse = await generateAiClient.generateText(generateTextRequest);
        console.log(JSON.stringify(generateTextResponse.generateTextResult, null, 2));
        return res.json(generateTextResponse.generateTextResult);
        */
    return "testresult";
  } catch (e) {
    console.log("error: ", e);
    throw e;
  }
}

// Read the Excel file
try {
  const workbook = xlsx.readFile(filePath);
  const [sheetName, cellRange] = cellRangeToProcess.split("!");
  const worksheet = workbook.Sheets[sheetName];

  // Extract column and row from the cell range
  const [startCell, endCell] = cellRange.split(":");
  const startRow = parseInt(startCell.match(/\d+/)[0], 10);
  const endRow = parseInt(endCell.match(/\d+/)[0], 10);
  const startCol = startCell.match(/[A-Z]+/)[0];
  const endCol = endCell.match(/[A-Z]+/)[0];

  // Loop through the specified cell range
  for (let row = startRow; row <= endRow; row++) {
    for (let col = startCol.charCodeAt(0); col <= endCol.charCodeAt(0); col++) {
      const cellAddress = String.fromCharCode(col) + row;
      const nextCellAddress = String.fromCharCode(col + 1) + row;

      const cellValue = worksheet[cellAddress]
        ? worksheet[cellAddress].v
        : null;
      const generatedText = generateText(cellValue);

      // Store the result in the adjacent cell on the right
      worksheet[nextCellAddress] = { v: generatedText, t: "s" }; // 's' indicates the data type is a string
    }
  }

  // Save a copy of the Excel file with datetime in the filename
  const originalFilename = path.basename(filePath, path.extname(filePath));
  const originalDir = path.dirname(filePath); // Extracting the directory of the original file
  const datetime = new Date()
    .toISOString()
    .replace(/[:\-T]/g, "")
    .split(".")[0]; // YYYYMMDDHHmmss format
  const newFilename = `${originalFilename}${datetime}.xlsx`;

  // Combining the original directory with the new filename
  const newFilePath = path.join(originalDir, newFilename);

  xlsx.writeFile(workbook, newFilePath);
  console.log(
    `A copy of the file has been saved as ${newFilename} in the same directory as the original file.`
  );
} catch (error) {
  console.error("Error reading or writing the Excel file: ", error);
}

