# Ensure TLS 1.2 is used for secure connections
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Import the Exchange Online Management module
Write-Host "Importing Exchange Online Management Module..." -ForegroundColor Cyan
Import-Module ExchangeOnlineManagement

# Connect to Exchange Online (Uncomment the Get-Credential if manual login is required)
Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
# $UserCredential = Get-Credential  
Connect-ExchangeOnline #-UserPrincipalName $UserCredential.UserName -ShowProgress $true

# Retrieve all user and shared mailboxes
Write-Host "Retrieving all user and shared mailboxes..." -ForegroundColor Cyan
$mailboxes = Get-EXOMailbox -ResultSize Unlimited | Where-Object { $_.RecipientTypeDetails -in @('UserMailbox', 'SharedMailbox') }

# Prepare an array to store all delegation data for exporting
$exportData = @()

foreach ($mailbox in $mailboxes) {
    Write-Host "`nProcessing mailbox: $($mailbox.PrimarySmtpAddress)..." -ForegroundColor Cyan

    # Collect Full Access Permissions
    $fullAccessPermissions = Get-EXOMailboxPermission -Identity $mailbox.PrimarySmtpAddress | 
        Where-Object { ($_.AccessRights -like "*FullAccess*") -and ($_.IsInherited -eq $false) }

    # Collect SendAs Permissions
    $sendAsPermissions = Get-EXORecipientPermission -Identity $mailbox.PrimarySmtpAddress | 
        Where-Object { ($_.AccessRights -like "*SendAs*") -and ($_.IsInherited -eq $false) }

    # Collect SendOnBehalf Permissions
    $sendOnBehalfPermissions = Get-EXOMailbox -Identity $mailbox.PrimarySmtpAddress | 
        Select-Object -ExpandProperty GrantSendOnBehalfTo -ErrorAction SilentlyContinue

    # Process Full Access permissions
    if ($fullAccessPermissions) {
        foreach ($permission in $fullAccessPermissions) {
            $exportData += [PSCustomObject]@{
                'Mailbox'    = $mailbox.PrimarySmtpAddress
                'User'       = $permission.User
                'AccessType' = "FullAccess"
            }
        }
    }

    # Process SendAs permissions
    if ($sendAsPermissions) {
        foreach ($permission in $sendAsPermissions) {
            $exportData += [PSCustomObject]@{
                'Mailbox'    = $mailbox.PrimarySmtpAddress
                'User'       = $permission.Trustee
                'AccessType' = "SendAs"
            }
        }
    }

    # Process SendOnBehalf permissions
    if ($sendOnBehalfPermissions) {
        foreach ($permission in $sendOnBehalfPermissions) {
            $exportData += [PSCustomObject]@{
                'Mailbox'    = $mailbox.PrimarySmtpAddress
                'User'       = $permission.Name
                'AccessType' = "SendOnBehalf"
            }
        }
    }

    Write-Host "Finished processing mailbox: $($mailbox.PrimarySmtpAddress)." -ForegroundColor Green
}

# Export the collected data to a CSV file
$outputFile = "C:\DelegationAccessSummary.csv"
Write-Host "Exporting delegation access data to $outputFile..." -ForegroundColor Cyan
$exportData | Export-Csv -Path $outputFile -NoTypeInformation -Force

# Disconnect from Exchange Online
Write-Host "Disconnecting from Exchange Online..." -ForegroundColor Cyan
Disconnect-ExchangeOnline -Confirm:$false

# End of script
Write-Host "Script execution completed. Data exported to $outputFile." -ForegroundColor Cyan
