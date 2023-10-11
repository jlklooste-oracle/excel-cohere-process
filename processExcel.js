const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');

// Check if file path is provided
if (process.argv.length <= 2) {
    console.log("Please provide the Excel file path as an argument.");
    process.exit(1);
}

const filePath = process.argv[2];

const ENDPOINT = "https://generativeai.aiservice.us-chicago-1.oci.oraclecloud.com"
const COMPARTMENT_ID = "ocid1.compartment.oc1..aaaaaaaaxdgp4qx2uodttcglrsg7b24vkuydba7hbsh6l3meaovxfw4ceybq"
const MODEL_ID = "cohere.command"
const MAX_TOKENS = 100
const STOP_SEQUENCE = "Input:"
const TEMPERATURE = 0

function generateText(prompt) {
    try {
        console.log("/generateText start")
        generateAiClient.endpoint = ENDPOINT;
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
            console.log("adding stopSequences", STOP_SEQUENCE)
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
        const generateTextResponse = await generateAiClient.generateText(generateTextRequest);
        console.log(JSON.stringify(generateTextResponse.generateTextResult, null, 2));
        return res.json(generateTextResponse.generateTextResult);
        */
        return "testresult"
    } catch (e) {
        console.log("error: ", e);
        throw e;
    }
});

// Read the Excel file
try {
    const workbook = xlsx.readFile(filePath);
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];

    // Read values from the specified cells
    const promptsPerMinute = 15
    const cellRangeToProcess = "Pursuit Analysis!AK82:AK82"    
    //const cellRangeToProcess = worksheet['B10'] ? worksheet['B10'].v : null;

    console.log({
        promptsPerMinute,
        cellRangeToProcess,
    });

    // Save a copy of the Excel file with datetime in the filename
    const originalFilename = path.basename(filePath, path.extname(filePath));
    const datetime = new Date().toISOString().replace(/[:\-T]/g, '').split('.')[0]; // YYYYMMDDHHmmss format
    const newFilename = `${originalFilename}${datetime}.xlsx`;

    xlsx.writeFile(workbook, newFilename);
    console.log(`A copy of the file has been saved as ${newFilename}`);
} catch (error) {
    console.error("Error reading or writing the Excel file: ", error);
}

//modify the code:
//fix errors. keep the "testresult" dummy response for now.
//change it so that cellRangeToProcess must come from a command line parameter
//add a validation that checks that the format is similar to this: "Worksheet Name!AK82:AK82"
//loop through the cell range
//  call the generateText for each of these cells, passing as a parameter the value of the cell
//  store the result of the function (in this case still a dummy result) in the cell that's on the right hand side of the cell
