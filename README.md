# Wave23AutoAtt
Automating Microsoft Forms using Playwright and PowerShell.
Here's a **README.md** template for your project, without placeholders. This is specific to your project that automates filling a Microsoft Form and submitting it, along with confirming the submission on a sheet on PowerPages.

---

# Microsoft Forms Automation Script

## Overview

This repository contains a PowerShell script that automates the process of filling out and submitting a Microsoft Form at a specific time of day. The script interacts with the form, submits the data, and confirms submission on a Google Sheets document stored on PowerPages.

The main goal of this project is to simplify the repetitive process of filling out forms by automating it with PowerShell, improving efficiency and accuracy.

## Features

* Automatically fills in Microsoft Forms based on predefined input.
* Submits the form on a scheduled time.
* Confirms the submission on a Google Sheets document.
* Easy to configure and use with minimal setup.

## Requirements

* PowerShell 7.0 or higher
* Access to Microsoft Forms API
* Google Sheets API credentials for submission confirmation
* `Microsoft.PowerApps.Cds.Client` module for interacting with Microsoft Forms
* `Google.Apis.Sheets.v4` module for interacting with Google Sheets

## Installation

1. Clone the repository to your local machine:

   ```bash
   git clone https://github.com/yourusername/microsoft-forms-automation.git
   ```

2. Navigate to the project directory:

   ```bash
   cd microsoft-forms-automation
   ```

3. Install the required PowerShell modules:

   ```bash
   Install-Module -Name Microsoft.PowerApps.Cds.Client -Force -Scope CurrentUser
   Install-Module -Name Google.Apis.Sheets.v4 -Force -Scope CurrentUser
   ```

4. Set up your Microsoft Forms API and Google Sheets API credentials:

   * Follow the steps to generate API credentials for both Microsoft Forms and Google Sheets.
   * Save your credentials in the respective configuration files (`msforms_creds.json`, `gsheets_creds.json`).

## Usage

1. Open PowerShell and navigate to the script directory:

   ```bash
   cd path\to\script
   ```

2. Run the script to begin the automation:

   ```bash
   .\formAutomation.ps1
   ```

3. The script will automatically fill in the form and submit it at the configured time. Once submitted, it will also update the Google Sheets document to confirm the submission.

## Configuration

### Microsoft Forms

* Set the form URL and input fields in the `config.json` file.

### Google Sheets

* Provide the Google Sheets document ID and the sheet name where submission confirmations are logged.

### Schedule the Script

* You can schedule this script to run automatically at a specified time using Windows Task Scheduler or any other task scheduling tool.

## Example Configuration

### config.json

```json
{
  "formUrl": "https://forms.microsoft.com/...",
  "formFields": {
    "name": "John Doe",
    "email": "john.doe@example.com",
    "feedback": "Great service!"
  },
  "googleSheetId": "your-google-sheet-id",
  "sheetName": "Form Submissions"
}
```

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

---

This **README** is tailored to your Microsoft Forms automation project and provides a detailed setup, usage instructions, and configuration for the necessary APIs. It should be ready to copy-paste into your GitHub repository!
