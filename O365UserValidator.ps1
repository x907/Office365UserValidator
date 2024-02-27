<#
.SYNOPSIS
This script validates a list of users against Office 365 and exports the valid users to a CSV file.

.DESCRIPTION
The script connects to Exchange Online, reads a CSV file containing a list of user recipients, and checks each user against Office 365. If a user is found, their details (name, email, and user principal name) are added to an array of valid users. The script then exports the valid users to a CSV file and disconnects from Exchange Online.

.PARAMETER csvPath
The path to the input CSV file containing the list of user recipients.

.PARAMETER outputCsvPath
The path to the output CSV file where the valid users will be exported.

.EXAMPLE
.\O365UserValidator.ps1 -csvPath "C:\temp\input.csv" -outputCsvPath "C:\temp\output.csv"
This example runs the script using the specified input CSV file and exports the valid users to the specified output CSV file.

.NOTES
- This script requires the ExchangeOnlineManagement module to be installed.
- Make sure to provide valid credentials to connect to Exchange Online.
#>


# Hardcoded variables
$csvPath = "C:\temp\input.csv"
$outputCsvPath = "C:\temp\output.csv"

# Connect to Exchange Online
try {
        Connect-ExchangeOnline
} catch {
    Write-Error "Failed to connect to Exchange Online. Please check your credentials."
    return
}

# Read the CSV file
$users = Import-Csv -Path $csvPath

# Prepare an array to hold the valid users
$validUsers = @()

    # Check each user
    foreach ($user in $users.Recipients) {
        try {
            # Attempt to get the user from Office 365
            $o365User = Get-Mailbox -Identity $user -ErrorAction Stop

            # If the above command didn't throw an error, the user is valid
            # Add their details to the array
            $validUser = New-Object PSObject
            $validUser | Add-Member -NotePropertyName Name -NotePropertyValue $o365User.DisplayName
            $validUser | Add-Member -NotePropertyName Email -NotePropertyValue $o365User.PrimarySmtpAddress
            $validUser | Add-Member -NotePropertyName UserPrincipalName -NotePropertyValue $o365User.UserPrincipalName

            $validUsers += $validUser
        } catch {
            # If an error occurred, the user wasn't found; ignore and continue
            Write-Error "User $user not found."
        }
    }

    # Export the valid users to a CSV
    $validUsers | Export-Csv -Path $outputCsvPath -NoTypeInformation

    # Disconnect from Exchange Online
    Disconnect-ExchangeOnline -Confirm:$false

    Write-Output "Process completed. Valid users exported to $outputCsvPath"
