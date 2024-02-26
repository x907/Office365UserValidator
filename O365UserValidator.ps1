<#
.SYNOPSIS
    Retrieves and validates Office 365 users based on a CSV file.

.DESCRIPTION
    The Get-O365Users function connects to Exchange Online, reads a CSV file containing user recipients,
    and checks each user to see if they exist in Office 365. The function then exports the valid users
    to a new CSV file.

.PARAMETER csvPath
    Specifies the path to the input CSV file containing user recipients.

.PARAMETER outputCsvPath
    Specifies the path to the output CSV file where the valid users will be exported.

.EXAMPLE
    Get-O365Users -csvPath "C:\Users\John\Documents\users.csv" -outputCsvPath "C:\Users\John\Documents\valid_users.csv"
    Retrieves and validates Office 365 users based on the "users.csv" file and exports the valid users to "valid_users.csv".

.NOTES
    This function requires the Exchange Online PowerShell module to be installed and the user to have the necessary permissions to connect to Exchange Online.
#>
function Get-O365Users {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [ValidateScript({Test-Path $_ -PathType Leaf})]
        [string]$csvPath,

        [Parameter(Mandatory=$true)]
        [ValidateScript({Test-Path $_ -PathType Container})]
        [string]$outputCsvPath
    )

# Connect to Exchange Online
try {
    $credential = Get-Credential
    Connect-ExchangeOnline -Credential $credential -ShowProgress $true
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
            # Continue adding other properties...

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
}