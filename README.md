# ServiceNow History Tracker

## Overview
This project automates interactions with the ServiceNow platform, focusing on retrieving historical records and processing specific events based on certain criteria. Using **Playwright**, this script automates browser interactions, clicks on history elements, and extracts relevant information for analysis.

## Technologies Used
- **Node.js**: Backend environment for running the automation script.
- **Playwright**: Browser automation framework used to interact with the ServiceNow platform.
- **XPath**: For navigating and selecting specific elements in the DOM.

## How It Works
1. **Navigate to the ServiceNow page**: The script starts by accessing the specified URL.
2. **Search for history elements**: The script searches for a specific history record (e.g., `INC0668760`) and clicks on it.
3. **Filter results by keywords**: It checks if the history record contains specific keywords ("Central de Atendimento" and "Fornecedor SZ").
4. **Data extraction**: When a matching record is found, the script extracts the date, time, and event details.
5. **Convert date to Excel format**: The extracted date is converted into a format compatible with Excel.

## Prerequisites
To run the script locally, make sure you have the following items installed:
- **Node.js** (v14 or higher)
- **Playwright**:
  ```bash
  npm install playwright
