
function Ensure-EXOSession {
    [CmdletBinding()]
    param()

    try {
        # Check if we're already connected (EXO session open)
        $isConnected = (Get-Module ExchangeOnlineManagement -ListAvailable) -and 
                       (Get-PSSession | Where-Object { $_.Name -like "ExchangeOnline*" })

        if (-not $isConnected) {
            Write-Host "Connecting to Exchange Online with app-only authentication."
            # Connect-ExchangeOnline -AppId $AppId -CertificateThumbprint $Thumbprint -Organization $Organization -ErrorAction Stop
            Connect-ExchangeOnline -UserPrincipalName "admin-aon@tremco-illbruck.com" -ErrorAction Stop
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
        Enable-Mailbox -Identity $UPN -Archive
        Write-Host "Archive enabled for: $UPN"

        Write-Host "Setting Retention Policy to '$RetentionPolicy' for mailbox: $UPN"
        Set-Mailbox -Identity $UPN -RetentionPolicy $RetentionPolicy
        Write-Host "Retention Policy updated successfully."
    }
    catch {
        Write-Error "$UPN Failed to enable archive or set retention for $($_.Exception.Message)"
    }
}

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
    }
}

$upn = "bruce.wayne@tremcocpg.com"

Ensure-EXOSession
# Invoke-ManagedFolderAssistantMaintenance -UPN $upn
Enable-MailboxArchiveAndRetention -UPN $upn -RetentionPolicy "9ec0ce2b-bdf7-4da3-9177-6c3ace6c4c8a"
