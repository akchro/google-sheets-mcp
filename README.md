# Google Sheets MCP
Model Context Protocol (MCP) integration with Google Sheets

## Features
This integration allows the LLM to 
- List all the spreadsheets of the user
- Create a new spreadsheet
- Copy a spreadsheet to another
- Write and Edit cells
- Fill tool for cells

## Get Started
1. [Create a new Google Cloud project](https://console.cloud.google.com/projectcreate)
2. [Enable the Google Drive API](https://console.cloud.google.com/workspace-api/products)
3. [Configure an OAuth consent screen](https://console.cloud.google.com/apis/credentials/consent) ("internal" is fine for testing)
4. [Create an OAuth Client ID](https://console.cloud.google.com/apis/credentials/oauthclient) for application type "Desktop App"
5. Download the JSON file of your client's OAuth keys
6. Rename the key file to `credentials.json` and place into the root of this repo (i.e. `google-sheets-mcp/credentials.json`)

> Note: If setting the user type to external, make sure to add your email for testing