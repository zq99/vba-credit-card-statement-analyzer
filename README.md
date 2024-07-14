
# VBA Excel Credit Card/Bank Statement Analyzer

## Overview

This VBA Excel application allows users to import a CSV file of credit card or bank transactions, and then generate basic pivot table reports for analysis. The application provides an easy-to-use interface with a button on the "Import" sheet that automates the import process and creates dynamic pivot table reports.

## Features

- **CSV Import:** Easily import CSV files containing credit card or bank statement data.
- **Automated Pivot Table Reports:** Automatically generate pivot table reports to analyze transaction data.
- **User-Friendly Interface:** Simple and intuitive button-click functionality for importing and analyzing data.

## Requirements

- Microsoft Excel (with support for VBA)
- A CSV file containing credit card or bank statement data

## Installation

1. **Open the Workbook:**
   - Open the `CreditCardStatementAnalyzer.xlsm` file in Microsoft Excel.

## Usage

1. **Open the `CreditCardStatementAnalyzer.xlsm` Workbook:**
   - Make sure macros are enabled. You may need to adjust your Excel security settings to allow macros to run.

2. **Importing the CSV File:**
   - Navigate to the `Import` sheet in the workbook.
   - Click the `Import` button located on this sheet.
   - Select the CSV file containing your credit card or bank statement data.
   - The data will be imported into the `Data` sheet.

3. **Viewing the Pivot Table Reports:**
   - After importing the CSV file, the application will automatically create pivot table reports.
   - The pivot tables will be available on the `Import` sheet for your analysis.

4. **Example File:**
   - The repository includes a dummy CSV file named `example_statement.csv` (generated using ChatGPT) so you can see how the application/code works.

## License

This project is open source and available under the [MIT License](LICENSE).
