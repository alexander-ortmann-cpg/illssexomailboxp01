param($Timer)

# ==========================================================
# Environment Variables / Configuration
# ==========================================================
$AppId = [System.Environment]::GetEnvironmentVariable("AppId", "Process")
$Organization = [System.Environment]::GetEnvironmentVariable("Organization", "Process")
$Thumbprint = [System.Environment]::GetEnvironmentVariable("Thumbprint", "Process")
$sqlConnectionString = [System.Environment]::GetEnvironmentVariable("sqlConnectionString", "Process")
$sendGridApiKey = [System.Environment]::GetEnvironmentVariable("SendGridApiKey", "Process")

# Logging – get the current universal time
$currentUTCtime = (Get-Date).ToUniversalTime()
Write-Host "=========================================================="
Write-Host "PowerShell Timer Trigger Function Starting..."
Write-Host "Current UTC Time: $currentUTCtime"
Write-Host "=========================================================="

if ($Timer.IsPastDue) {
    Write-Warning "PowerShell timer is running late!"
}

# Example HTML body content
$bodyHtml = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="utf-8" />
    <title>Mailbox Archive Enabled</title>
    <style>
        /* Example custom list style */
        .step-list {
            list-style: decimal inside;
            margin-top: 10px;
            margin-bottom: 10px;
            padding-left: 1rem;
        }
        .step-list li {
            margin-bottom: 8px; /* spacing between steps */
        }
    </style>
</head>
<body style=""background-color: #f8f9fa;"">
    <div class=""container my-5"">
        <div class=""card shadow"">
            <div class=""card-header bg-danger text-white"">
                <h3>Important: Archive Mailbox Enabled</h3>
            </div>
            <div class=""card-body"">
                <p class=""lead"">Your mailbox has exceeded the critical storage threshold.</p>
                <p>
                    To ensure you can continue sending and receiving emails, an Archive Mailbox has been 
                    automatically enabled for your account.
                </p>
                
                <p>
                    <strong>How It Works with Outlook:</strong>
                </p>
                <ol class=""step-list"">
                    <li>Open your Outlook client and navigate to the folder list.</li>
                    <li>You should now see a new "Archive" mailbox.</li>
                    <li>Messages that are older or exceed your primary mailbox quota 
                        will be automatically moved into this archive folder, 
                        helping keep your main mailbox size manageable.</li>
                    <li>You can still search your archive the same way you search 
                        your primary mailbox.</li>
                </ol>
                
                <p>
                    If you have any questions, please contact the help desk.
                </p>
            </div>
            <p class=""ms-3"">
                Thank you
            </p>
        </div>
    </div>

</body>
</html>
"@

# ==========================================================
# 1) Function: Ensure-EXOSession
#    - Connects to Exchange Online using App-Only Authentication (if not already connected)
# ==========================================================
function Ensure-EXOSession {
    [CmdletBinding()]
    param()

    try {
        # Check if we're already connected (EXO session open)
        $isConnected = (Get-Module ExchangeOnlineManagement -ListAvailable) -and 
                       (Get-PSSession | Where-Object { $_.Name -like "ExchangeOnline*" })

        if (-not $isConnected) {
            Write-Host "Connecting to Exchange Online with app-only authentication."
            Connect-ExchangeOnline -AppId $AppId -CertificateThumbprint $Thumbprint -Organization $Organization -ErrorAction Stop
            # Connect-ExchangeOnline -UserPrincipalName "admin-aon@tremco-illbruck.com" -ErrorAction Stop
            Write-Host "Connected to Exchange Online."
        }
        else {
            Write-Host "Already connected to Exchange Online."
        }
    }
    catch {
        Write-Error "Failed to connect to Exchange Online: $($_.Exception.Message)"
        throw  # Rethrow so the calling scope can handle it
    }
}

# ==========================================================
# 2) Function: Get-BytesFromString
#    - Parses a string like "2.5 GB (2,684,354,560 bytes)" and returns the numeric bytes.
#    - If "Unlimited," returns 0 (or handle differently if needed).
# ==========================================================
function Get-BytesFromString {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SizeString
    )

    if ($SizeString -match 'Unlimited') {
        return 0
    }

    if ($SizeString -match '\(([\d,]+)\s+bytes\)') {
        $rawNumber = $Matches[1].Replace(',', '')  # remove commas
        return [Int64]$rawNumber
    }

    return 0
}

# ==========================================================
# 3) Function: Send-SendGridMail
#    - Sends an email via SendGrid API.
#    - Adjust or replace with your actual sending approach.
# ==========================================================
function Send-SendGridMail {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)] [string]$To,
        [Parameter(Mandatory = $true)] [string]$From,
        [Parameter(Mandatory = $true)] [string]$Subject,
        [Parameter(Mandatory = $true)] [string]$Body
    )

    try {
        if (-not $sendGridApiKey) {
            Write-Error "SendGrid API key not found in environment variable 'SendGridApiKey'."
            return
        }

        # Construct the JSON body for the SendGrid API
        $jsonBody = @{
            personalizations = @(
                @{
                    to = @(
                        @{ email = $To }
                    )
                }
            )
            from             = @{ email = $From }
            subject          = $Subject
            content          = @(
                @{
                    type  = "text/html"
                    value = $Body
                }
            )
        } | ConvertTo-Json -Depth 5

        # Set headers and send the request
        $headers = @{
            "Authorization" = "Bearer $sendGridApiKey"
            "Content-Type"  = "application/json"
        }

        Write-Host "Sending email via SendGrid to: $To"
        Invoke-RestMethod -Uri "https://api.sendgrid.com/v3/mail/send" `
            -Method Post `
            -Headers $headers `
            -Body $jsonBody
        Write-Host "Email sent successfully."
    }
    catch {
        Write-Error "Failed to send email via SendGrid: $($_.Exception.Message)"
    }
}

# ==========================================================
# 4) Function: Invoke-ManagedFolderAssistantMaintenance
#    - Runs Start-ManagedFolderAssistant multiple times for a mailbox UPN.
# ==========================================================
function Invoke-ManagedFolderAssistantMaintenance {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UPN
    )

    try {
        Write-Host "Starting Managed Folder Assistant Maintenance for mailbox: $UPN"
        
        # 1st invocation
        Start-ManagedFolderAssistant -Identity $UPN
        Write-Host "Start-ManagedFolderAssistant for $UPN completed (pass 1)."

        # 2nd invocation
        Start-ManagedFolderAssistant -Identity $UPN
        Write-Host "Start-ManagedFolderAssistant for $UPN completed (pass 2)."

        # FullCrawl
        Start-ManagedFolderAssistant -Identity $UPN -FullCrawl
        Write-Host "Start-ManagedFolderAssistant with -FullCrawl for $UPN completed."

        # HoldCleanup
        Start-ManagedFolderAssistant -Identity $UPN -HoldCleanup
        Write-Host "Start-ManagedFolderAssistant with -HoldCleanup for $UPN completed."

    }
    catch {
        Write-Error "Failed to run ManagedFolderAssistant for $UPN : $($_.Exception.Message)"
        # Log the error to Azure SQL Error Log
        Log-ErrorToAzureSql -UserPrincipalName $UPN -ErrorMessage "ManagedFolderAssistant failed: $($_.Exception.Message)" -SqlConnectionString $sqlConnectionString
    }
}

# ==========================================================
# 5) Function: Enable-MailboxArchiveAndRetention
#    - Enables the mailbox archive and updates the retention policy.
#    - Returns $true if successful, $false otherwise.
# ==========================================================
function Enable-MailboxArchiveAndRetention {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UPN,

        [Parameter(Mandatory = $true)]
        [string]$RetentionPolicy
    )

    try {
        Write-Host "Enabling mailbox archive for: $UPN"
        Enable-Mailbox -Identity $UPN -Archive -ErrorAction Stop
        Write-Host "Archive enabled for: $UPN"

        Write-Host "Setting Retention Policy to '$RetentionPolicy' for mailbox: $UPN"
        Set-Mailbox -Identity $UPN -RetentionPolicy $RetentionPolicy -ErrorAction Stop
        Write-Host "Retention Policy updated successfully."

        return $true
    }
    catch {
        Write-Error "$UPN failed to enable archive or set retention: $($_.Exception.Message)"
        # Log the error to Azure SQL Error Log
        Log-ErrorToAzureSql -UserPrincipalName $UPN -ErrorMessage "Enable-MailboxArchiveAndRetention failed: $($_.Exception.Message)" -SqlConnectionString $sqlConnectionString
        return $false
    }
}

# =====================================================================
# Function: Log-MailboxUsageToAzureSql
# Purpose:  Inserts the mailbox usage data (over threshold entries) into
#           an Azure SQL Database table for logging purposes.
#
# Note: Ensure that your table (dbo.MailboxUsageLog) has been updated to
# include an additional column (e.g. ErrorMessage) if you wish to log errors.
# For example, the table could have these columns:
#   UserPrincipalName, ObjectID, PrimaryMailboxUsed, RecoverableItemsUsed,
#   PrimaryMailboxQuota, RecoverableItemsQuota, [Log], ErrorMessage
# =====================================================================
function Log-MailboxUsageToAzureSql {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Collections.ArrayList] $ReportEntries,

        [Parameter(Mandatory = $true)]
        [string] $SqlConnectionString,

        [Parameter(Mandatory = $false)]
        [string] $LogMessage = "Mailbox usage exceeds threshold"
    )

    try {
        # Create and open the connection
        $connection = New-Object System.Data.SqlClient.SqlConnection($SqlConnectionString)
        $connection.Open()

        foreach ($entry in $ReportEntries) {
            $sqlCommandText = @"
INSERT INTO dbo.MailboxUsageLog (
    UserPrincipalName,
    ObjectID,
    PrimaryMailboxUsed,
    RecoverableItemsUsed,
    PrimaryMailboxQuota,
    RecoverableItemsQuota,
    [Log],
    ErrorMessage
)
VALUES (
    @UserPrincipalName,
    @ObjectID,
    @PrimaryMailboxUsed,
    @RecoverableItemsUsed,
    @PrimaryMailboxQuota,
    @RecoverableItemsQuota,
    @LogMessage,
    @ErrorMessage
);
"@

            $command = $connection.CreateCommand()
            $command.CommandText = $sqlCommandText

            $null = $command.Parameters.AddWithValue("@UserPrincipalName", $entry.UserPrincipalName)
            $null = $command.Parameters.AddWithValue("@ObjectID", $entry.ObjectID)
            $null = $command.Parameters.AddWithValue("@PrimaryMailboxUsed", $entry."PrimaryMailboxUsed(%)")
            $null = $command.Parameters.AddWithValue("@RecoverableItemsUsed", $entry."RecoverableItemsUsed(%)")
            $null = $command.Parameters.AddWithValue("@PrimaryMailboxQuota", $entry."PrimaryMailboxQuota")
            $null = $command.Parameters.AddWithValue("@RecoverableItemsQuota", $entry."RecoverableItemsQuota")
            $null = $command.Parameters.AddWithValue("@LogMessage", $LogMessage)
            # Pass an error message if present; if not, use an empty string
            $null = $command.Parameters.AddWithValue("@ErrorMessage", ($entry.ErrorMessage -as [string]) )

            $command.ExecuteNonQuery() | Out-Null
        }

        $connection.Close()
        Write-Host "Successfully logged $($ReportEntries.Count) mailbox entries into Azure SQL."
    }
    catch {
        Write-Error "Error writing mailbox usage data to Azure SQL: $($_.Exception.Message)"
    }
}

# ==========================================================
# NEW FUNCTION: Log-ErrorToAzureSql
# Purpose:  Logs errors to a dedicated Azure SQL table (e.g. dbo.MailboxErrorLog).
#
# Ensure your SQL table has the following columns (or similar):
#   - UserPrincipalName (NVARCHAR)
#   - ErrorMessage (NVARCHAR(MAX))
#   - ErrorTimestamp (DATETIME) -- can be set using GETUTCDATE()
# ==========================================================
function Log-ErrorToAzureSql {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserPrincipalName,
        [Parameter(Mandatory = $true)]
        [string]$ErrorMessage,
        [Parameter(Mandatory = $true)]
        [string]$SqlConnectionString
    )

    try {
        $connection = New-Object System.Data.SqlClient.SqlConnection($SqlConnectionString)
        $connection.Open()
        
        $sqlCommandText = @"
INSERT INTO dbo.MailboxErrorLog (
    UserPrincipalName,
    ErrorMessage,
    ErrorTimestamp
)
VALUES (
    @UserPrincipalName,
    @ErrorMessage,
    GETUTCDATE()
);
"@
        $command = $connection.CreateCommand()
        $command.CommandText = $sqlCommandText

        $null = $command.Parameters.AddWithValue("@UserPrincipalName", $UserPrincipalName)
        $null = $command.Parameters.AddWithValue("@ErrorMessage", $ErrorMessage)

        $command.ExecuteNonQuery() | Out-Null
        $connection.Close()
        Write-Host "Error for $UserPrincipalName logged successfully in Azure SQL."
    }
    catch {
        Write-Error "Error logging to Azure SQL Error Log: $($_.Exception.Message)"
    }
}

# ==========================================================
# Main Processing Logic
# ==========================================================
try {
    # Connect to Exchange Online
    Ensure-EXOSession
    
    # Build a report object for usage threshold events
    $report = [System.Collections.ArrayList]::new()

    Write-Host "Retrieving user mailboxes..."
    # Replace the following line with your desired mailbox query:
    $mailboxes = Get-Mailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited -ErrorAction Stop
    # $mailboxes = Get-Mailbox "msteele@tremcoinc.com" -RecipientTypeDetails UserMailbox -ErrorAction Stop

    foreach ($mailbox in $mailboxes) {
        $upn = $mailbox.UserPrincipalName

        # Retrieve mailbox statistics (with error handling)
        $mbStats = $null
        try {
            $mbStats = Get-MailboxStatistics -Identity $upn -ErrorAction Stop
        }
        catch {
            $errorMsg = "Get-MailboxStatistics failed for $upn $($_.Exception.Message). Trying to reconnect..."
            Write-Warning $errorMsg

            # Log the error for this mailbox
            Log-ErrorToAzureSql -UserPrincipalName $upn -ErrorMessage $errorMsg -SqlConnectionString $sqlConnectionString

            Disconnect-ExchangeOnline -Confirm:$false
            Ensure-EXOSession

            $mbStats = Get-MailboxStatistics -Identity $upn -ErrorAction SilentlyContinue
            if (-not $mbStats) {
                $errorMsg2 = "Could not retrieve mailbox statistics for $upn even after reconnect. Skipping..."
                Write-Warning $errorMsg2
                Log-ErrorToAzureSql -UserPrincipalName $upn -ErrorMessage $errorMsg2 -SqlConnectionString $sqlConnectionString
                continue
            }
        }

        # Primary mailbox size and quota
        $primarySizeString = $mbStats.TotalItemSize.ToString()
        $mailboxSizeBytes = Get-BytesFromString -SizeString $primarySizeString

        $mailboxQuotaString = $mailbox.ProhibitSendQuota.ToString()
        $mailboxQuotaBytes = if ($mailboxQuotaString -match "unlimited") { 0 } else { 
            Get-BytesFromString -SizeString $mailboxQuotaString 
        }
        
        [double]$mailboxUsagePercent = 0
        if ($mailboxQuotaBytes -gt 0) {
            $mailboxUsagePercent = [math]::Round(($mailboxSizeBytes / $mailboxQuotaBytes) * 100, 2)
        }

        # Recoverable Items size and quota
        $recoverableSizeString = $mbStats.TotalDeletedItemSize.ToString()
        $recoverableSizeBytes = Get-BytesFromString -SizeString $recoverableSizeString

        $recoverableQuotaString = $mailbox.RecoverableItemsQuota.ToString()
        $recoverableQuotaBytes = if ($recoverableQuotaString -match "unlimited") { 0 } else {
            Get-BytesFromString -SizeString $recoverableQuotaString
        }

        [double]$recoverableUsagePercent = 0
        if ($recoverableQuotaBytes -gt 0) {
            $recoverableUsagePercent = [math]::Round(($recoverableSizeBytes / $recoverableQuotaBytes) * 100, 2)
        }

        Write-Host "Processing: $upn"
        Write-Host "  Primary mailbox size (bytes): $mailboxSizeBytes ($mailboxUsagePercent%)"
        Write-Host "  Recoverable items size (bytes): $recoverableSizeBytes ($recoverableUsagePercent%)"

        # If primary mailbox usage exceeds threshold and no archive exists, enable archive and retention, then send email on success
        if (($mailboxUsagePercent -gt 95) -and ($mailbox.ArchiveStatus -ne "Active")) {
            Write-Warning "Mailbox Usage Percent for $upn is over 95% usage and no archive is present!"
            $result = Enable-MailboxArchiveAndRetention -UPN $upn -RetentionPolicy "9ec0ce2b-bdf7-4da3-9177-6c3ace6c4c8a"
            if ($result) {
                # Uncomment and adjust the following to send an email notification if desired:
                Send-SendGridMail -To $upn `
                    -From "no-reply@tremcocpg.com" `
                    -Subject "Archive Mailbox Enabled for Your Account" `
                    -Body $bodyHtml
            }
            else {
                $errMsg = "Failed to enable archive and retention for $upn, email not sent."
                Write-Error $errMsg
                Log-ErrorToAzureSql -UserPrincipalName $upn -ErrorMessage $errMsg -SqlConnectionString $sqlConnectionString
            }
        }

        # If recoverable items usage exceeds threshold, run Managed Folder Assistant maintenance
        if ($recoverableUsagePercent -gt 95) {
            Write-Warning "Recoverable Items Usage Percent for $upn is over 95% usage!"
            Invoke-ManagedFolderAssistantMaintenance -UPN $upn
        }

        # Check Archive Usage for mailboxes with an active archive.
        if ($mailbox.ArchiveStatus -eq "Active") {
            Write-Host "Mailbox $upn has an archive enabled. Checking archive usage..."
            $archiveStats = Get-MailboxStatistics -Identity $upn -Archive -ErrorAction SilentlyContinue
            if ($archiveStats) {
                $archiveSizeString = $archiveStats.TotalItemSize.ToString()
                $archiveSizeBytes = Get-BytesFromString -SizeString $archiveSizeString

                $archiveQuotaString = $mailbox.ArchiveQuota.ToString()
                $archiveQuotaBytes = if ($archiveQuotaString -match "unlimited") { 0 } else {
                    Get-BytesFromString -SizeString $archiveQuotaString
                }

                [double]$archiveUsagePercent = 0
                if ($archiveQuotaBytes -gt 0) {
                    $archiveUsagePercent = [math]::Round(($archiveSizeBytes / $archiveQuotaBytes) * 100, 2)
                }

                Write-Host "Processing Archive for: $upn"
                Write-Host "  Archive size (bytes): $archiveSizeBytes ($archiveUsagePercent%)"
                if ($archiveUsagePercent -gt 95) {
                    Write-Warning "Archive usage for $upn is over 95%. Enabling auto-expanding archive."
                    try {
                        Set-Mailbox -Identity $upn -AutoExpandingArchive $true -ErrorAction Stop
                        Write-Host "Auto-expanding archive enabled for $upn."
                    }
                    catch {
                        $errMsgArchive = "Failed to enable auto-expanding archive for $upn $($_.Exception.Message)"
                        Write-Error $errMsgArchive
                        Log-ErrorToAzureSql -UserPrincipalName $upn -ErrorMessage $errMsgArchive -SqlConnectionString $sqlConnectionString
                    }
                }
            }
            else {
                $errMsgArchiveStats = "Could not retrieve archive statistics for $upn."
                Write-Warning $errMsgArchiveStats
                Log-ErrorToAzureSql -UserPrincipalName $upn -ErrorMessage $errMsgArchiveStats -SqlConnectionString $sqlConnectionString
            }
        }

        # Log the mailbox if either primary or recoverable usage is over threshold
        if ($mailboxUsagePercent -gt 95 -or $recoverableUsagePercent -gt 95) {
            Write-Warning "Mailbox $upn is over 95% usage!"
            # Note: If there was an error during processing for this mailbox, you can add an ErrorMessage property.
            $report += [PSCustomObject]@{
                UserPrincipalName         = $upn
                ObjectID                  = $mailbox.ExternalDirectoryObjectId
                "PrimaryMailboxUsed(%)"   = $mailboxUsagePercent
                "RecoverableItemsUsed(%)" = $recoverableUsagePercent
                "PrimaryMailboxQuota"     = $mailboxQuotaString
                "RecoverableItemsQuota"   = $recoverableQuotaString
                ErrorMessage              = ""  # or set to an error message if one was captured
            }
        }
    }

    if ($report.Count -gt 0) {
        Log-MailboxUsageToAzureSql -ReportEntries $report `
            -SqlConnectionString $sqlConnectionString `
            -LogMessage "Mailbox usage threshold exceeded"
        
        Write-Host "=========================================================="
        Write-Host "Generating usage report. Found $($report.Count) mailboxes over threshold."
        $report | Format-Table -AutoSize
    }
    else {
        Write-Host "No mailboxes exceeding threshold. No report generated."
    }
}
catch {
    Write-Error "An error occurred in the main block: $($_.Exception.Message)"
    Write-Error "Stack Trace: $($_.ScriptStackTrace)"
    # Optionally log a global error (using a generic UPN or context identifier)
    Log-ErrorToAzureSql -UserPrincipalName "Global" -ErrorMessage "Main block error: $($_.Exception.Message)" -SqlConnectionString $sqlConnectionString
}
finally {
    Write-Host "Disconnecting from Exchange Online..."
    Disconnect-ExchangeOnline -Confirm:$false
    Write-Host "Script execution finished."
}
