# Connect to Office 365
$UserCredential = Get-Credential
Connect-ExchangeOnline -Credential $UserCredential

# Path to the CSV file containing the email addresses
$csvPath = "C:\path\to\your\file.csv"

# Path where the output CSV will be saved
$outputCsvPath = "C:\path\to\your\outputFile.csv"

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
        $validUsers += [PSCustomObject]@{
            Name               = $o365User.DisplayName
            Email              = $o365User.PrimarySmtpAddress
            UserPrincipalName  = $o365User.UserPrincipalName
            Department         = $o365User.Department
        }
    } catch {
        # If an error occurred, the user wasn't found; ignore and continue
        Write-Output "User $user not found."
    }
}

# Export the valid users to a CSV
$validUsers | Export-Csv -Path $outputCsvPath -NoTypeInformation

# Disconnect from Exchange Online
Disconnect-ExchangeOnline -Confirm:$false

Write-Output "Process completed. Valid users exported to $outputCsvPath"
