# ---------------------------------------------------------
# Connect to Exchange Online
# ---------------------------------------------------------
Import-Module ExchangeOnlineManagement

# Connect (adjust your admin UPN as needed). 
# If -UseRPSSession is forced to $true, objects become deserialized.
Connect-ExchangeOnline -UserPrincipalName "<YourAdmin@tenant.onmicrosoft.com>" -ShowBanner:$false

# ---------------------------------------------------------
# Helper function to parse "X GB (Y bytes)" strings
# ---------------------------------------------------------

function Ensure-EXOSession {
    param (
        [string]$AdminUPN
    )

    # Check if we already have a connected session
    $isConnected = (Get-Module ExchangeOnlineManagement -ListAvailable) -and (Get-PSSession | Where-Object { $_.Name -like "ExchangeOnline*" })
    
    if (-not $isConnected) {
        Connect-ExchangeOnline -UserPrincipalName $AdminUPN -ErrorAction Stop
    }
}

function Get-BytesFromString {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SizeString
    )

    # If string says "Unlimited", return 0 or handle however you like
    if ($SizeString -match 'Unlimited') {
        return 0
    }

    # Look for something like "(2,718,185,472 bytes)" using regex
    # Group ([\d,]+) captures digits plus commas
    # Then remove commas before converting to [Int64]
    if ($SizeString -match '\(([\d,]+)\s+bytes\)') {
        $rawNumber = $Matches[1].Replace(',', '')  # strip out commas
        return [Int64]$rawNumber
    }

    # If it doesn't match, default to 0
    return 0
}

$AdminUPN = "admin-aon@tremco-illbruck.com"
Ensure-EXOSession -AdminUPN $AdminUPN

# ---------------------------------------------------------
# Retrieve Mailboxes and Generate Report
# ---------------------------------------------------------
$report = @()

try {
    # Check if the Exchange Online session is connected
    Ensure-EXOSession -AdminUPN $AdminUPN

    $null = Get-Mailbox -ErrorAction Stop

    # Get only user mailboxes (excluding shared/room if not desired)
    Get-Mailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited | ForEach-Object {
        $mailbox = $_

        # Use a unique property (PrimarySmtpAddress)
        $mbStats = Get-MailboxStatistics -Identity $mailbox.UserPrincipalName -ErrorAction SilentlyContinue

        # Ensure we actually got results
        if ($null -eq $mbStats) {
            Write-Warning "Could not retrieve mailbox statistics for $($mailbox.PrimarySmtpAddress). Skipping..."
            return
        }
        
        # --- Primary Mailbox Size ---
        # Deserialized object is usually a string like "2.5 GB (2,684,354,560 bytes)"
        $primarySizeString = $mbStats.TotalItemSize.ToString()
        $mailboxSizeBytes = Get-BytesFromString -SizeString $primarySizeString

        # Fetch the ProhibitSendQuota from the mailbox (could also be "Unlimited")
        $mailboxQuotaString = (Get-Mailbox $mailbox.Identity).ProhibitSendQuota.ToString()
        # Parse the numeric bytes from that string
        $mailboxQuotaBytes = 0
        if ($mailboxQuotaString -notmatch "unlimited") {
            $mailboxQuotaBytes = Get-BytesFromString -SizeString $mailboxQuotaString
        }

        # Calculate usage percentage (avoid division by zero)
        $mailboxUsagePercent = 0
        if ($mailboxQuotaBytes -gt 0) {
            $mailboxUsagePercent = [math]::Round(($mailboxSizeBytes / $mailboxQuotaBytes) * 100, 2)
        }

        # --- Recoverable Items Size ---
        $recoverableSizeString = $mbStats.TotalDeletedItemSize.ToString()
        $recoverableSizeBytes = Get-BytesFromString -SizeString $recoverableSizeString

        # Fetch the RecoverableItemsQuota (could be unlimited)
        $recoverableQuotaString = (Get-Mailbox $mailbox.Identity).RecoverableItemsQuota.ToString()
        $recoverableQuotaBytes = 0
        if ($recoverableQuotaString -notmatch "unlimited") {
            $recoverableQuotaBytes = Get-BytesFromString -SizeString $recoverableQuotaString
        }

        # Calculate usage percentage (avoid division by zero)
        $recoverableUsagePercent = 0
        if ($recoverableQuotaBytes -gt 0) {
            $recoverableUsagePercent = [math]::Round(($recoverableSizeBytes / $recoverableQuotaBytes) * 100, 2)
        }

        Write-Host "Processing $($mailbox.UserPrincipalName)..."
        Write-Host "Primary Mailbox: $mailboxSizeBytes bytes ($mailboxUsagePercent%)"

        # --- Check if either usage is over 90% ---
        if ($mailboxUsagePercent -gt 90 -or $recoverableUsagePercent -gt 90) {
            Write-Host "WARNING: $($mailbox.UserPrincipalName) is over 90% usage!"
            $report += [PSCustomObject]@{
                UserPrincipalName         = $mailbox.UserPrincipalName
                ObjectID                  = $mailbox.ExternalDirectoryObjectId
                "PrimaryMailboxUsed(%)"   = $mailboxUsagePercent
                "RecoverableItemsUsed(%)" = $recoverableUsagePercent
                "PrimaryMailboxQuota"     = $mailboxQuotaString
                "RecoverableItemsQuota"   = $recoverableQuotaString
            }
        }
    }

}
catch {
    Write-Error "Could not connect to Exchange Online. Exiting..."
    Write-Host "Caught error: $_"
    Write-Host "Stack trace: $($_.ScriptStackTrace)"
    Write-Host "Invocation info: $($_.InvocationInfo)"
}

# ---------------------------------------------------------
# Output the report
# ---------------------------------------------------------
$report | Format-Table -AutoSize

$csvPath = "C:\Temp\MailboxUsageReport.csv"
$report | Export-Csv -Path $csvPath -NoTypeInformation

# If you want to export it:
# $report | Export-Csv -Path "C:\Temp\MailboxUsageReport.csv" -NoTypeInformation

# Disconnect from Exchange Online
Disconnect-ExchangeOnline -Confirm:$false
