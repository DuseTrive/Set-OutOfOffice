<#
.SYNOPSIS
    Set Out of Office for multiple mailboxes using Exchange Online
.DESCRIPTION
    Interactive script to configure automatic replies for multiple mailboxes
.NOTES
    Requires: Exchange Online Management PowerShell module
#>

# Ensure required modules are installed
function Test-RequiredModules {
    $requiredModule = 'ExchangeOnlineManagement'
    
    if (-not (Get-Module -ListAvailable -Name $requiredModule)) {
        Write-Host "Installing module: $requiredModule" -ForegroundColor Yellow
        Install-Module -Name $requiredModule -Scope CurrentUser -Force -AllowClobber
    }
}

# Function to clean and parse mailbox list
function Get-CleanMailboxList {
    param([string]$RawInput)
    
    $mailboxes = $RawInput -split "`n" | ForEach-Object {
        $_.Trim() -replace '\s+', ''
    } | Where-Object { 
        $_ -match '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$' 
    }
    
    return $mailboxes | Select-Object -Unique
}

# Function to convert plain text message to HTML for Exchange
function ConvertTo-HtmlMessage {
    param([string]$PlainText)
    
    # Replace line breaks with HTML breaks
    $htmlMessage = $PlainText -replace "`r`n", "<br>" -replace "`n", "<br>"
    
    # Wrap in basic HTML structure
    $htmlMessage = @"
<html>
<body>
$htmlMessage
</body>
</html>
"@
    
    return $htmlMessage
}

# Function to parse date and time
function Get-DateTime {
    param(
        [string]$Prompt
    )
    
    do {
        Write-Host "`n$Prompt" -ForegroundColor Cyan
        Write-Host "Format: DD/MM/YYYY HH:MM AM/PM (e.g., 25/12/2024 09:00 AM)" -ForegroundColor Gray
        $input = Read-Host "Enter date and time"
        
        try {
            $dateTime = [DateTime]::ParseExact($input, 'dd/MM/yyyy hh:mm tt', $null)
            $valid = $true
        }
        catch {
            Write-Host "Invalid format. Please try again." -ForegroundColor Red
            $valid = $false
        }
    } while (-not $valid)
    
    return $dateTime
}

# Function to set OOO for a single mailbox
function Set-MailboxOOO {
    param(
        [string]$Mailbox,
        [DateTime]$StartTime,
        [DateTime]$EndTime,
        [string]$Message
    )
    
    try {
        # Convert message to HTML format for proper line break display
        $htmlMessage = ConvertTo-HtmlMessage -PlainText $Message
        
        # Set automatic replies using Exchange Online
        Set-MailboxAutoReplyConfiguration -Identity $Mailbox `
                                        -AutoReplyState Scheduled `
                                        -StartTime $StartTime `
                                        -EndTime $EndTime `
                                        -InternalMessage $htmlMessage `
                                        -ExternalMessage $htmlMessage `
                                        -ExternalAudience All `
                                        -ErrorAction Stop
        
        Write-Host "[SUCCESS] Set OOO for: $Mailbox" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "[FAILED] Could not set OOO for: $Mailbox" -ForegroundColor Red
        Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# Main Script
Clear-Host
Write-Host "=========================================================" -ForegroundColor Cyan
Write-Host "     Out of Office Configuration Script" -ForegroundColor Cyan
Write-Host "=========================================================" -ForegroundColor Cyan

# Check and install required modules
Write-Host "`nChecking required modules..." -ForegroundColor Yellow
Test-RequiredModules

# Connect to Exchange Online
Write-Host "`nConnecting to Exchange Online..." -ForegroundColor Yellow
try {
    # Clean up any existing Exchange sessions
    Get-PSSession | Where-Object {$_.ConfigurationName -eq "Microsoft.Exchange"} | Remove-PSSession -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
    
    # Connect to Exchange Online
    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
    $null = Get-OrganizationConfig -ErrorAction Stop
    Write-Host "[SUCCESS] Exchange Online connected successfully" -ForegroundColor Green
}
catch {
    Write-Host "[INFO] Initial Exchange connection failed, retrying..." -ForegroundColor Yellow
    try {
        if ($env:AUTOMATED_EXECUTION -ne 'true') {
            $exchCred = Get-Credential -Message "Enter Exchange Online Admin Credentials"
            Connect-ExchangeOnline -Credential $exchCred -ShowBanner:$false -ErrorAction Stop
        } else {
            # For automated execution, try without credentials
            Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        }
        $null = Get-OrganizationConfig -ErrorAction Stop
        Write-Host "[SUCCESS] Exchange Online connected successfully" -ForegroundColor Green
    }
    catch {
        if ($env:AUTOMATED_EXECUTION -eq 'true') {
            Write-Host "[WARNING] Exchange connection failed in automated mode - continuing" -ForegroundColor Yellow
        } else {
            Write-Host "[FAILED] Exchange Online connection failed after retry: $_" -ForegroundColor Red
            exit
        }
    }
}

# Step 1: Get mailbox list
Write-Host "`n---------------------------------------------------------" -ForegroundColor Gray
Write-Host "STEP 1: Mailbox List" -ForegroundColor Cyan
Write-Host "---------------------------------------------------------" -ForegroundColor Gray
Write-Host "Paste the list of email addresses (press Enter twice when done):" -ForegroundColor Yellow

$rawMailboxInput = @()
do {
    $line = Read-Host
    if ($line) {
        $rawMailboxInput += $line
    }
} while ($line)

$mailboxList = Get-CleanMailboxList -RawInput ($rawMailboxInput -join "`n")

if ($mailboxList.Count -eq 0) {
    Write-Host "`n[ERROR] No valid email addresses found!" -ForegroundColor Red
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
    exit
}

# Step 2: Get start date/time
Write-Host "`n---------------------------------------------------------" -ForegroundColor Gray
Write-Host "STEP 2: Schedule Duration" -ForegroundColor Cyan
Write-Host "---------------------------------------------------------" -ForegroundColor Gray
$startDateTime = Get-DateTime -Prompt "Enter START date and time"

# Step 3: Get end date/time
$endDateTime = Get-DateTime -Prompt "Enter END date and time"

# Validate date range
if ($endDateTime -le $startDateTime) {
    Write-Host "`n[ERROR] End date/time must be after start date/time!" -ForegroundColor Red
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
    exit
}

# Step 4: Get OOO message
Write-Host "`n---------------------------------------------------------" -ForegroundColor Gray
Write-Host "STEP 3: Out of Office Message" -ForegroundColor Cyan
Write-Host "---------------------------------------------------------" -ForegroundColor Gray
Write-Host "Enter the Out of Office message (press Enter twice when done):" -ForegroundColor Yellow
Write-Host "Note: Line breaks and formatting will be preserved in the email." -ForegroundColor Gray

$oooMessageLines = @()
do {
    $line = Read-Host
    if ($line -or $oooMessageLines.Count -gt 0) {
        $oooMessageLines += $line
    }
} while ($line -or $oooMessageLines.Count -eq 0)

$oooMessage = ($oooMessageLines | Select-Object -SkipLast 1) -join "`n"

# Step 5: Display configuration for confirmation
Write-Host "`n=========================================================" -ForegroundColor Cyan
Write-Host "CONFIGURATION SUMMARY" -ForegroundColor Cyan
Write-Host "=========================================================" -ForegroundColor Cyan

Write-Host "`nMailboxes ($($mailboxList.Count) total):" -ForegroundColor Yellow
$mailboxList | ForEach-Object { Write-Host "  - $_" -ForegroundColor White }

Write-Host "`nSchedule:" -ForegroundColor Yellow
Write-Host "  From: $($startDateTime.ToString('dd/MM/yyyy hh:mm tt'))" -ForegroundColor White
Write-Host "  To:   $($endDateTime.ToString('dd/MM/yyyy hh:mm tt'))" -ForegroundColor White
Write-Host "  Duration: $([Math]::Round(($endDateTime - $startDateTime).TotalDays, 1)) days" -ForegroundColor White

Write-Host "`nOut of Office Message:" -ForegroundColor Yellow
Write-Host "---------------------------------------------------------" -ForegroundColor Gray
$oooMessage -split "`n" | ForEach-Object {
    Write-Host "  $_" -ForegroundColor White
}
Write-Host "---------------------------------------------------------" -ForegroundColor Gray

# Confirm to proceed
Write-Host "`n=========================================================" -ForegroundColor Cyan
$confirm = Read-Host "Do you want to proceed? (Y/N)"
if ($confirm -notmatch '^[Yy]') {
    Write-Host "`n[CANCELLED] Operation cancelled by user" -ForegroundColor Yellow
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
    exit
}

# Step 6: Test with first mailbox
Write-Host "`n=========================================================" -ForegroundColor Cyan
Write-Host "TESTING WITH FIRST MAILBOX" -ForegroundColor Cyan
Write-Host "=========================================================" -ForegroundColor Cyan

$testMailbox = $mailboxList[0]
Write-Host "`nProcessing test mailbox: $testMailbox" -ForegroundColor Yellow

$testResult = Set-MailboxOOO -Mailbox $testMailbox -StartTime $startDateTime -EndTime $endDateTime -Message $oooMessage

if (-not $testResult) {
    Write-Host "`n[FAILED] Test failed. Please check the error and try again." -ForegroundColor Red
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
    exit
}

# Confirm to proceed with remaining mailboxes
Write-Host "`n---------------------------------------------------------" -ForegroundColor Gray
Write-Host "Please verify the Out of Office settings for: $testMailbox" -ForegroundColor Yellow
$confirmRemaining = Read-Host "`nDo you want to proceed with the remaining $($mailboxList.Count - 1) mailbox(es)? (Y/N)"

if ($confirmRemaining -notmatch '^[Yy]') {
    Write-Host "`n[STOPPED] Operation stopped after test mailbox" -ForegroundColor Yellow
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
    exit
}

# Step 7: Process remaining mailboxes
if ($mailboxList.Count -gt 1) {
    Write-Host "`n=========================================================" -ForegroundColor Cyan
    Write-Host "PROCESSING REMAINING MAILBOXES" -ForegroundColor Cyan
    Write-Host "=========================================================" -ForegroundColor Cyan
    
    $remainingMailboxes = $mailboxList | Select-Object -Skip 1
    $successCount = 1  # Already processed first one
    $failCount = 0
    
    foreach ($mailbox in $remainingMailboxes) {
        Write-Host "`nProcessing: $mailbox" -ForegroundColor Yellow
        $result = Set-MailboxOOO -Mailbox $mailbox -StartTime $startDateTime -EndTime $endDateTime -Message $oooMessage
        
        if ($result) {
            $successCount++
        } else {
            $failCount++
        }
        
        Start-Sleep -Milliseconds 500
    }
    
    # Final summary
    Write-Host "`n=========================================================" -ForegroundColor Cyan
    Write-Host "OPERATION COMPLETE" -ForegroundColor Cyan
    Write-Host "=========================================================" -ForegroundColor Cyan
    Write-Host "[SUCCESS] Successfully configured: $successCount mailbox(es)" -ForegroundColor Green
    if ($failCount -gt 0) {
        Write-Host "[FAILED] Failed: $failCount mailbox(es)" -ForegroundColor Red
    }
}

# Disconnect from Exchange Online
Write-Host "`nDisconnecting from Exchange Online..." -ForegroundColor Yellow
try {
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
    Write-Host "[SUCCESS] Disconnected from Exchange Online" -ForegroundColor Green
}
catch {
    Write-Host "[WARNING] Could not disconnect from Exchange Online" -ForegroundColor Yellow
}

Write-Host "`n=========================================================" -ForegroundColor Cyan
Write-Host "Script completed successfully!" -ForegroundColor Green
Write-Host "=========================================================" -ForegroundColor Cyan
