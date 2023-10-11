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

// Retrieve the basePrompt from prompt.txt
let basePrompt;
try {
  basePrompt = fs.readFileSync("prompt.txt", "utf8");
} catch (e) {
  console.error("Error reading the prompt.txt file:", e);
  process.exit(1);
}

//this is used to wait between requests in order to comply with the rate limit
async function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function generateText(prompt) {
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
    return "<result><question1>no</question1><question2>yes</question2><question3>no</question3><question4>no</question4><question5>no</question5><question6>no</question6><question7>no</question7><question8>no</question8><question9>no</question9><question10>no</question10><question11>no</question11><question12>no</question12><question13>no</question13></result>";
  } catch (e) {
    console.log("error: ", e);
    throw e;
  }
}

function colToNum(col) {
  let num = 0;
  for (let i = 0; i < col.length; i++) {
    num = num * 26 + col.charCodeAt(i) - 64;
  }
  return num;
}

function numToCol(num) {
  let col = "";
  while (num > 0) {
    const modulo = (num - 1) % 26;
    col = String.fromCharCode(65 + modulo) + col;
    num = Math.floor((num - modulo) / 26);
  }
  return col;
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
  const col = startCell.match(/[A-Z]+/)[0];
  console.log("col", col);

  for (let row = startRow; row <= endRow; row++) {
    const cellAddress = col + row;
    console.log("cellAddress", cellAddress);

    // Convert the column label to a number, increment it, and convert back to a label
    const nextCol = numToCol(colToNum(col) + 1);
    const nextCellAddress = nextCol + row;
    const cellValue = worksheet[cellAddress] ? worksheet[cellAddress].v : null;

    if (!worksheet[nextCellAddress] || !worksheet[nextCellAddress].v) {
      const startTime = Date.now();
      const fullPrompt = basePrompt + cellValue + "\r\nOutput: ";
      const generatedText = await generateText(fullPrompt); // Call generateText with full text
      worksheet[nextCellAddress] = { v: generatedText, t: "s" };

      const endTime = Date.now();
      const timeTaken = endTime - startTime;
      const timeToWait = MS_BETWEEN_REQUESTS - timeTaken;

      if (timeToWait > 0) {
        await sleep(timeToWait);
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

//your task is:
//we're trying to set nextCellAddress to the column to the right of the current column (col). how can we do that? keep in mind that the column could also use two letters, e.g. AK.
