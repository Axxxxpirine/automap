# automap

A small Python script that calculates driving distance and travel time between
a single origin address and a list of destination addresses stored in an Excel file,
using the openrouteservice API.

## Requirements

Install dependencies:

\`\`\`bash
pip install -r requirements.txt
\`\`\`

## Input file

The script expects a file called \`addresses.xlsx\` with the following columns:

| Address | PostalCode | City |
|--------|------------|------|

## Environment variable

Set your openrouteservice API key:

macOS / Linux:
\`\`\`bash
export OPENROUTESERVICE_API_KEY="your_api_key"
\`\`\`

Windows (PowerShell):
\`\`\`powershell
setx OPENROUTESERVICE_API_KEY "your_api_key"
\`\`\`

## Run

From the project folder:

\`\`\`bash
python automap.py
\`\`\`

The script will read \`addresses.xlsx\` and create \`addresses_with_distances.xlsx\`.
