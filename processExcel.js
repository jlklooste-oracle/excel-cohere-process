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
const MS_BETWEEN_REQUESTS = 61000 / MAX_RPM;

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

//this is used to wait between requests in order to comply with the rate limit
async function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
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

async function processCells() {
  // Read the Excel file
  const workbook = xlsx.readFile(filePath);
  const [sheetName, cellRange] = cellRangeToProcess.split("!");
  const worksheet = workbook.Sheets[sheetName];

  // Extract column and row from the cell range
  const [startCell, endCell] = cellRange.split(":");
  const startRow = parseInt(startCell.match(/\d+/)[0], 10);
  const endRow = parseInt(endCell.match(/\d+/)[0], 10);
  const startCol = startCell.match(/[A-Z]+/)[0];
  const endCol = endCell.match(/[A-Z]+/)[0];

  for (let row = startRow; row <= endRow; row++) {
    for (let col = startCol.charCodeAt(0); col <= endCol.charCodeAt(0); col++) {
      const cellAddress = String.fromCharCode(col) + row;
      const nextCellAddress = String.fromCharCode(col + 1) + row;
      const cellValue = worksheet[cellAddress]
        ? worksheet[cellAddress].v
        : null;

      if (!worksheet[nextCellAddress] || !worksheet[nextCellAddress].v) {
        const startTime = Date.now();
        const generatedText = await generateText(cellValue);
        worksheet[nextCellAddress] = { v: generatedText, t: "s" };

        const endTime = Date.now();
        const timeTaken = endTime - startTime;
        const timeToWait = MS_BETWEEN_REQUESTS - timeTaken;

        if (timeToWait > 0) {
          await sleep(timeToWait);
        }
      }
    }
  }

  // Save a copy of the Excel file with datetime in the filename
  const originalFilename = path.basename(filePath, path.extname(filePath));
  const originalDir = path.dirname(filePath);
  const datetime = new Date()
    .toISOString()
    .replace(/[:\-T]/g, "")
    .split(".")[0];
  const newFilename = `${originalFilename}${datetime}.xlsx`;
  const newFilePath = path.join(originalDir, newFilename);

  xlsx.writeFile(workbook, newFilePath);
  console.log(
    `A copy of the file has been saved as ${newFilename} in the same directory as the original file.`
  );
}

// Execute the process
try {
  processCells();
} catch (error) {
  console.error("Error reading or writing the Excel file: ", error);
}
