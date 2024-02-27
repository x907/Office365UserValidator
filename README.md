# Office 365 User Validator

This PowerShell script validates a list of users against Office 365 and exports the valid users to a CSV file. It's a handy tool for administrators who need to verify the existence of a large number of users in Office 365.

![PowerShell Logo](https://upload.wikimedia.org/wikipedia/commons/2/2f/PowerShell_5.0_icon.png)

## Features

- Validates users against Office 365
- Exports valid users to a CSV file
- Provides detailed error messages for invalid users

## Prerequisites

- PowerShell 5.1 or higher
- ExchangeOnlineManagement module

## Usage

1. Update the `$csvPath` and `$outputCsvPath` variables in the script with your input and output file paths respectively.

2. Run the script in PowerShell:

    ```
    .\O365UserValidator.ps1
    ```

## Parameters

- `csvPath`: The path to the input CSV file containing the list of user recipients.
- `outputCsvPath`: The path to the output CSV file where the valid users will be exported.

## Example

    ```
    .\O365UserValidator.ps1 -csvPath "C:\temp\input.csv" -outputCsvPath "C:\temp\output.csv"
    ```

This example runs the script using the specified input CSV file and exports the valid users to the specified output CSV file.

## Notes

- This script requires the ExchangeOnlineManagement module to be installed.
- Make sure to provide valid credentials to connect to Exchange Online.

## License

This project is dedicated to the public domain as per the [Unlicense](http://unlicense.org/).