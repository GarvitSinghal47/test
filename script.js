const { google } = require("googleapis");

const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
const SHEET_NAME = process.env.SHEET_NAME;
const GOOGLE_CREDENTIALS = process.env.GOOGLE_CREDENTIALS;

// Validate environment variables
function validateEnvVariables() {
  const missingVars = [];
  if (!SPREADSHEET_ID) missingVars.push("SPREADSHEET_ID");
  if (!SHEET_NAME) missingVars.push("SHEET_NAME");
  if (!GOOGLE_CREDENTIALS) missingVars.push("GOOGLE_CREDENTIALS");

  if (missingVars.length > 0) {
    throw new Error(
      `Missing required environment variables: ${missingVars.join(", ")}`
    );
  }
}

// Set up GoogleAuth for Google Sheets API
async function authorize() {
  try {
    const credentials = JSON.parse(GOOGLE_CREDENTIALS);
    const auth = new google.auth.GoogleAuth({
      credentials,
      scopes: ["https://www.googleapis.com/auth/spreadsheets"],
    });
    const authClient = await auth.getClient();
    return google.sheets({ version: "v4", auth: authClient });
  } catch (error) {
    throw new Error(
      `Failed to authorize: ${error.message}. Please check your GOOGLE_CREDENTIALS.`
    );
  }
}

// Update Google Sheets with pull request data
async function updateSpreadsheet(prData) {
  const sheets = await authorize();

  try {
    // Fetch existing rows from the sheet
    const { data } = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: SHEET_NAME,
    });

    const existingRows = data.values || [];
    let rowToUpdate = null;

    // Find the row that corresponds to the pull request URL
    for (let i = 1; i < existingRows.length; i++) {
      if (existingRows[i][1] === prData[1]) { // prData[1] is the PR URL
        rowToUpdate = i + 1; // Get the row number to update
        break;
      }
    }

    if (rowToUpdate) {
      // Check if row data has changed and update if necessary
      const existingData = existingRows[rowToUpdate - 1];
      const prDataString = JSON.stringify(prData);
      const existingDataString = JSON.stringify(existingData);

      if (prDataString !== existingDataString) {
        console.log(`Detected changes for row ${rowToUpdate}.`);

        const updates = [];
        const columns = ["A", "B", "C", "D", "E", "F", "G", "H"];
        for (let col = 0; col < prData.length; col++) {
          if (existingData[col] !== prData[col]) {
            updates.push({
              range: `${SHEET_NAME}!${columns[col]}${rowToUpdate}`,
              values: [[prData[col]]],
            });
          }
        }

        if (updates.length > 0) {
          // Batch update changed columns
          await sheets.spreadsheets.values.batchUpdate({
            spreadsheetId: SPREADSHEET_ID,
            resource: {
              data: updates,
              valueInputOption: "RAW",
            },
          });
          console.log(`Updated row ${rowToUpdate} in Google Sheets.`);
        } else {
          console.log(`No changes detected for row ${rowToUpdate}.`);
        }
      } else {
        console.log(`No changes detected for row ${rowToUpdate}.`);
      }
    } else {
      // Append new row
      await sheets.spreadsheets.values.append({
        spreadsheetId: SPREADSHEET_ID,
        range: `${SHEET_NAME}!A:H`,
        valueInputOption: "RAW",
        resource: { values: [prData] },
      });
      console.log(`Added new row to Google Sheets.`);
    }
  } catch (error) {
    throw new Error(`Failed to update spreadsheet: ${error.message}`);
  }
}

// Main function to handle pull request changes
async function handlePullRequestChange(prData) {
  // Filter out PRs from members of the organization
  if (
    prData[8] !== 'true' && // user_site_admin
    prData[9] === 'User' && // user_type
    !prData[10].includes("MEMBER") // author_association
  ) {
    await updateSpreadsheet(prData);
  } else {
    console.log(
      "PR skipped: Author is a member of the organization or a site admin."
    );
  }
}

// Validate environment variables
try {
  validateEnvVariables();
} catch (error) {
  console.error(error.message);
  process.exit(1);
}

// Parse command-line arguments
const prData = process.argv.slice(2);

if (prData.length !== 11) {
  console.error("Incorrect number of arguments provided.");
  process.exit(1);
}

// Run the script
handlePullRequestChange(prData).catch((error) => {
  console.error("An error occurred:", error.message);
  process.exit(1);
});
