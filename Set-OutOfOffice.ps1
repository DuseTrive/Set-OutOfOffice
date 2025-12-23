<#
.SYNOPSIS
    Set Out of Office for multiple mailboxes using Microsoft Graph
.DESCRIPTION
    Interactive script to configure automatic replies for multiple mailboxes
.NOTES
    Requires: Microsoft.Graph PowerShell modules and Exchange Online Management
#>

# Ensure required modules are installed
function Test-RequiredModules {
    $requiredModules = @(
        'Microsoft.Graph.Users.Actions',
        'Microsoft.Graph.Authentication',
        'ExchangeOnlineManagement'
    )
    
    foreach ($module in $requiredModules) {
        if (-not (Get-Module -ListAvailable -Name $module)) {
            Write-Host "Installing module: $module" -ForegroundColor Yellow
            Install-Module -Name $module -Scope CurrentUser -Force -AllowClobber
        }
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
        # Convert to ISO 8601 format for Graph API
        $startTimeISO = $StartTime.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ss.fffZ')
        $endTimeISO = $EndTime.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ss.fffZ')
        
        # Create automatic replies settings
        $automaticRepliesSetting = @{
            status = "scheduled"
            scheduledStartDateTime = @{
                dateTime = $startTimeISO
                timeZone = "UTC"
            }
            scheduledEndDateTime = @{
                dateTime = $endTimeISO
                timeZone = "UTC"
            }
            internalReplyMessage = $Message
            externalReplyMessage = $Message
        }
        
        # Set automatic replies using Graph API
        Update-MgUserMailboxSetting -UserId $Mailbox -AutomaticRepliesSetting $automaticRepliesSetting
        
        Write-Host "✓ Successfully set OOO for: $Mailbox" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "✗ Failed to set OOO for: $Mailbox" -ForegroundColor Red
        Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# Main Script
Clear-Host
Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "     Out of Office Configuration Script" -ForegroundColor Cyan
Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan

# Check and install required modules
Write-Host "`nChecking required modules..." -ForegroundColor Yellow
Test-RequiredModules

# Connect to Microsoft Graph
Write-Host "`nConnecting to Microsoft Graph..." -ForegroundColor Yellow
try {
    Connect-MgGraph -Scopes "MailboxSettings.ReadWrite", "User.ReadWrite.All" -NoWelcome
    Write-Host "✓ Connected to Microsoft Graph" -ForegroundColor Green
}
catch {
    Write-Host "✗ Failed to connect to Microsoft Graph" -ForegroundColor Red
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    exit
}

# Step 1: Get mailbox list
Write-Host "`n" + ("─" * 55) -ForegroundColor Gray
Write-Host "STEP 1: Mailbox List" -ForegroundColor Cyan
Write-Host ("─" * 55) -ForegroundColor Gray
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
    Write-Host "`n✗ No valid email addresses found!" -ForegroundColor Red
    exit
}

# Step 2: Get start date/time
Write-Host "`n" + ("─" * 55) -ForegroundColor Gray
Write-Host "STEP 2: Schedule Duration" -ForegroundColor Cyan
Write-Host ("─" * 55) -ForegroundColor Gray
$startDateTime = Get-DateTime -Prompt "Enter START date and time"

# Step 3: Get end date/time
$endDateTime = Get-DateTime -Prompt "Enter END date and time"

# Validate date range
if ($endDateTime -le $startDateTime) {
    Write-Host "`n✗ End date/time must be after start date/time!" -ForegroundColor Red
    exit
}

# Step 4: Get OOO message
Write-Host "`n" + ("─" * 55) -ForegroundColor Gray
Write-Host "STEP 3: Out of Office Message" -ForegroundColor Cyan
Write-Host ("─" * 55) -ForegroundColor Gray
Write-Host "Enter the Out of Office message (press Enter twice when done):" -ForegroundColor Yellow

$oooMessageLines = @()
do {
    $line = Read-Host
    if ($line -or $oooMessageLines.Count -gt 0) {
        $oooMessageLines += $line
    }
} while ($line -or $oooMessageLines.Count -eq 0)

$oooMessage = ($oooMessageLines | Select-Object -SkipLast 1) -join "`n"

# Step 5: Display configuration for confirmation
Write-Host "`n" + ("═" * 55) -ForegroundColor Cyan
Write-Host "CONFIGURATION SUMMARY" -ForegroundColor Cyan
Write-Host ("═" * 55) -ForegroundColor Cyan

Write-Host "`nMailboxes ($($mailboxList.Count) total):" -ForegroundColor Yellow
$mailboxList | ForEach-Object { Write-Host "  • $_" -ForegroundColor White }

Write-Host "`nSchedule:" -ForegroundColor Yellow
Write-Host "  From: $($startDateTime.ToString('dd/MM/yyyy hh:mm tt'))" -ForegroundColor White
Write-Host "  To:   $($endDateTime.ToString('dd/MM/yyyy hh:mm tt'))" -ForegroundColor White
Write-Host "  Duration: $([Math]::Round(($endDateTime - $startDateTime).TotalDays, 1)) days" -ForegroundColor White

Write-Host "`nOut of Office Message:" -ForegroundColor Yellow
Write-Host "┌$("─" * 53)┐" -ForegroundColor Gray
$oooMessage -split "`n" | ForEach-Object {
    Write-Host "│ $($_.PadRight(52))│" -ForegroundColor White
}
Write-Host "└$("─" * 53)┘" -ForegroundColor Gray

# Confirm to proceed
Write-Host "`n" + ("═" * 55) -ForegroundColor Cyan
$confirm = Read-Host "Do you want to proceed? (Y/N)"
if ($confirm -notmatch '^[Yy]') {
    Write-Host "`n✗ Operation cancelled by user" -ForegroundColor Yellow
    Disconnect-MgGraph | Out-Null
    exit
}

# Step 6: Test with first mailbox
Write-Host "`n" + ("═" * 55) -ForegroundColor Cyan
Write-Host "TESTING WITH FIRST MAILBOX" -ForegroundColor Cyan
Write-Host ("═" * 55) -ForegroundColor Cyan

$testMailbox = $mailboxList[0]
Write-Host "`nProcessing test mailbox: $testMailbox" -ForegroundColor Yellow

$testResult = Set-MailboxOOO -Mailbox $testMailbox -StartTime $startDateTime -EndTime $endDateTime -Message $oooMessage

if (-not $testResult) {
    Write-Host "`n✗ Test failed. Please check the error and try again." -ForegroundColor Red
    Disconnect-MgGraph | Out-Null
    exit
}

# Confirm to proceed with remaining mailboxes
Write-Host "`n" + ("─" * 55) -ForegroundColor Gray
Write-Host "Please verify the Out of Office settings for: $testMailbox" -ForegroundColor Yellow
$confirmRemaining = Read-Host "`nDo you want to proceed with the remaining $($mailboxList.Count - 1) mailbox(es)? (Y/N)"

if ($confirmRemaining -notmatch '^[Yy]') {
    Write-Host "`n✗ Operation stopped after test mailbox" -ForegroundColor Yellow
    Disconnect-MgGraph | Out-Null
    exit
}

# Step 7: Process remaining mailboxes
if ($mailboxList.Count -gt 1) {
    Write-Host "`n" + ("═" * 55) -ForegroundColor Cyan
    Write-Host "PROCESSING REMAINING MAILBOXES" -ForegroundColor Cyan
    Write-Host ("═" * 55) -ForegroundColor Cyan
    
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
    Write-Host "`n" + ("═" * 55) -ForegroundColor Cyan
    Write-Host "OPERATION COMPLETE" -ForegroundColor Cyan
    Write-Host ("═" * 55) -ForegroundColor Cyan
    Write-Host "✓ Successfully configured: $successCount mailbox(es)" -ForegroundColor Green
    if ($failCount -gt 0) {
        Write-Host "✗ Failed: $failCount mailbox(es)" -ForegroundColor Red
    }
}

# Disconnect
Write-Host "`nDisconnecting from Microsoft Graph..." -ForegroundColor Yellow
Disconnect-MgGraph | Out-Null
Write-Host "✓ Disconnected" -ForegroundColor Green

Write-Host "`n" + ("═" * 55) -ForegroundColor Cyan
Write-Host "Script completed successfully!" -ForegroundColor Green
Write-Host ("═" * 55) -ForegroundColor Cyan