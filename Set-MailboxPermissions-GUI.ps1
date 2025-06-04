#region Description
<#     
.NOTES
==============================================================================
Created on:         2025/06/09
Created by:         Drago Petrovic
Organization:       MSB365.blog
Filename:           Set-MailboxPermissions-GUI.ps1
Current version:    V1.0     

Find us on:
* Website:         https://www.msb365.blog
* Technet:         https://social.technet.microsoft.com/Profile/MSB365
* LinkedIn:        https://www.linkedin.com/in/drago-petrovic/
* MVP Profile:     https://mvp.microsoft.com/de-de/PublicProfile/5003446
==============================================================================

.SYNOPSIS
    Manages Exchange Online mailbox permissions based on CSV file input with GUI file picker
.DESCRIPTION
    This script sets "Send As" and "Full Access" permissions for mailboxes defined in a CSV file.
    It removes any existing permissions that are not defined in the CSV file.
    Features a GUI file picker for easy CSV file selection.
.PARAMETER WhatIf
    Shows what would be done without making changes
.EXAMPLE
    .\Set-MailboxPermissions-GUI.ps1
.EXAMPLE
    .\Set-MailboxPermissions-GUI.ps1 -WhatIf

    .EXAMPLE CSV File Format
    MailboxIdentity,UserIdentity,SendAs,FullAccess
    shared.mailbox@company.com,john.doe@company.com,TRUE,TRUE
    shared.mailbox@company.com,jane.smith@company.com,TRUE,FALSE
    shared.mailbox@company.com,admin@company.com,FALSE,TRUE
    finance@company.com,john.doe@company.com,TRUE,TRUE
    finance@company.com,finance.manager@company.com,TRUE,TRUE

    The script will process each mailbox and apply the permissions as specified in the CSV file.
    It will also remove any permissions not listed in the CSV file, preserving system permissions.

.COPYRIGHT
Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), 
to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, 
and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, 
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, 
WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
===========================================================================
.CHANGE LOG
V1.00, 2025/06/09 - DrPe - Initial version



--- keep it simple, but significant ---


--- by MSB365 Blog ---

#>
#endregion
##############################################################################################################
[cmdletbinding()]
param(
[switch]$accepteula,
[switch]$v)

###############################################################################
#Script Name variable
$Scriptname = "Set MailboxPermissions EXO"
$RKEY = "MSB365_Set-MailboxPermissions-EXO"
###############################################################################

[void][System.Reflection.Assembly]::Load('System.Drawing, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
[void][System.Reflection.Assembly]::Load('System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')

function ShowEULAPopup($mode)
{
    $EULA = New-Object -TypeName System.Windows.Forms.Form
    $richTextBox1 = New-Object System.Windows.Forms.RichTextBox
    $btnAcknowledge = New-Object System.Windows.Forms.Button
    $btnCancel = New-Object System.Windows.Forms.Button

    $EULA.SuspendLayout()
    $EULA.Name = "MIT"
    $EULA.Text = "$Scriptname - License Agreement"

    $richTextBox1.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $richTextBox1.Location = New-Object System.Drawing.Point(12,12)
    $richTextBox1.Name = "richTextBox1"
    $richTextBox1.ScrollBars = [System.Windows.Forms.RichTextBoxScrollBars]::Vertical
    $richTextBox1.Size = New-Object System.Drawing.Size(776, 397)
    $richTextBox1.TabIndex = 0
    $richTextBox1.ReadOnly=$True
    $richTextBox1.Add_LinkClicked({Start-Process -FilePath $_.LinkText})
    $richTextBox1.Rtf = @"
{\rtf1\ansi\ansicpg1252\deff0\nouicompat{\fonttbl{\f0\fswiss\fprq2\fcharset0 Segoe UI;}{\f1\fnil\fcharset0 Calibri;}{\f2\fnil\fcharset0 Microsoft Sans Serif;}}
{\colortbl ;\red0\green0\blue255;}
{\*\generator Riched20 10.0.19041}{\*\mmathPr\mdispDef1\mwrapIndent1440 }\viewkind4\uc1
\pard\widctlpar\f0\fs19\lang1033 MSB365 SOFTWARE MIT LICENSE\par
Copyright (c) 2025 Drago Petrovic\par
$Scriptname \par
\par
{\pict{\*\picprop}\wmetafile8\picw26\pich26\picwgoal32000\pichgoal15
0100090000035000000000002700000000000400000003010800050000000b0200000000050000
000c0202000200030000001e000400000007010400040000000701040027000000410b2000cc00
010001000000000001000100000000002800000001000000010000000100010000000000000000
000000000000000000000000000000000000000000ffffff00000000ff040000002701ffff0300
00000000
}These license terms are an agreement between you and MSB365 (or one of its affiliates). IF YOU COMPLY WITH THESE LICENSE TERMS, YOU HAVE THE RIGHTS BELOW. BY USING THE SOFTWARE, YOU ACCEPT THESE TERMS.\par
\par
MIT License\par
{\pict{\*\picprop}\wmetafile8\picw26\pich26\picwgoal32000\pichgoal15
0100090000035000000000002700000000000400000003010800050000000b0200000000050000
000c0202000200030000001e000400000007010400040000000701040027000000410b2000cc00
010001000000000001000100000000002800000001000000010000000100010000000000000000
000000000000000000000000000000000000000000ffffff00000000ff040000002701ffff0300
00000000
}\par
\pard
{\pntext\f0 1.\tab}{\*\pn\pnlvlbody\pnf0\pnindent0\pnstart1\pndec{\pntxta.}}
\fi-360\li360 Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions: \par
\pard\widctlpar\par
\pard\widctlpar\li360 The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.\par
\par
\pard\widctlpar\fi-360\li360 2.\tab THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. \par
\pard\widctlpar\par
\pard\widctlpar\fi-360\li360 3.\tab IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE. \par
\pard\widctlpar\par
\pard\widctlpar\fi-360\li360 4.\tab DISCLAIMER OF WARRANTY. THE SOFTWARE IS PROVIDED \ldblquote AS IS,\rdblquote  WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL MSB365 OR ITS LICENSORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THE SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.\par
\pard\widctlpar\par
\pard\widctlpar\fi-360\li360\qj 5.\tab LIMITATION ON AND EXCLUSION OF DAMAGES. IF YOU HAVE ANY BASIS FOR RECOVERING DAMAGES DESPITE THE PRECEDING DISCLAIMER OF WARRANTY, YOU CAN RECOVER FROM MICROSOFT AND ITS SUPPLIERS ONLY DIRECT DAMAGES UP TO U.S. $1.00. YOU CANNOT RECOVER ANY OTHER DAMAGES, INCLUDING CONSEQUENTIAL, LOST PROFITS, SPECIAL, INDIRECT, OR INCIDENTAL DAMAGES. This limitation applies to (i) anything related to the Software, services, content (including code) on third party Internet sites, or third party applications; and (ii) claims for breach of contract, warranty, guarantee, or condition; strict liability, negligence, or other tort; or any other claim; in each case to the extent permitted by applicable law. It also applies even if MSB365 knew or should have known about the possibility of the damages. The above limitation or exclusion may not apply to you because your state, province, or country may not allow the exclusion or limitation of incidental, consequential, or other damages.\par
\pard\widctlpar\par
\pard\widctlpar\fi-360\li360\qj 6.\tab ENTIRE AGREEMENT. This agreement, and any other terms MSB365 may provide for supplements, updates, or third-party applications, is the entire agreement for the software.\par
\pard\widctlpar\qj\par
\pard\widctlpar\fi-360\li360\qj 7.\tab A complete script documentation can be found on the website https://www.msb365.blog.\par
\pard\widctlpar\par
\pard\sa200\sl276\slmult1\f1\fs22\lang9\par
\pard\f2\fs17\lang2057\par
}
"@
    $richTextBox1.BackColor = [System.Drawing.Color]::White
    $btnAcknowledge.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
    $btnAcknowledge.Location = New-Object System.Drawing.Point(544, 415)
    $btnAcknowledge.Name = "btnAcknowledge";
    $btnAcknowledge.Size = New-Object System.Drawing.Size(119, 23)
    $btnAcknowledge.TabIndex = 1
    $btnAcknowledge.Text = "Accept"
    $btnAcknowledge.UseVisualStyleBackColor = $True
    $btnAcknowledge.Add_Click({$EULA.DialogResult=[System.Windows.Forms.DialogResult]::Yes})

    $btnCancel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
    $btnCancel.Location = New-Object System.Drawing.Point(669, 415)
    $btnCancel.Name = "btnCancel"
    $btnCancel.Size = New-Object System.Drawing.Size(119, 23)
    $btnCancel.TabIndex = 2
    if($mode -ne 0)
    {
   $btnCancel.Text = "Close"
    }
    else
    {
   $btnCancel.Text = "Decline"
    }
    $btnCancel.UseVisualStyleBackColor = $True
    $btnCancel.Add_Click({$EULA.DialogResult=[System.Windows.Forms.DialogResult]::No})

    $EULA.AutoScaleDimensions = New-Object System.Drawing.SizeF(6.0, 13.0)
    $EULA.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Font
    $EULA.ClientSize = New-Object System.Drawing.Size(800, 450)
    $EULA.Controls.Add($btnCancel)
    $EULA.Controls.Add($richTextBox1)
    if($mode -ne 0)
    {
   $EULA.AcceptButton=$btnCancel
    }
    else
    {
        $EULA.Controls.Add($btnAcknowledge)
   $EULA.AcceptButton=$btnAcknowledge
        $EULA.CancelButton=$btnCancel
    }
    $EULA.ResumeLayout($false)
    $EULA.Size = New-Object System.Drawing.Size(800, 650)

    Return ($EULA.ShowDialog())
}

function ShowEULAIfNeeded($toolName, $mode)
{
$eulaRegPath = "HKCU:Software\Microsoft\$RKEY"
$eulaAccepted = "No"
$eulaValue = $toolName + " EULA Accepted"
if(Test-Path $eulaRegPath)
{
$eulaRegKey = Get-Item $eulaRegPath
$eulaAccepted = $eulaRegKey.GetValue($eulaValue, "No")
}
else
{
$eulaRegKey = New-Item $eulaRegPath
}
if($mode -eq 2) # silent accept
{
$eulaAccepted = "Yes"
        $ignore = New-ItemProperty -Path $eulaRegPath -Name $eulaValue -Value $eulaAccepted -PropertyType String -Force
}
else
{
if($eulaAccepted -eq "No")
{
$eulaAccepted = ShowEULAPopup($mode)
if($eulaAccepted -eq [System.Windows.Forms.DialogResult]::Yes)
{
        $eulaAccepted = "Yes"
        $ignore = New-ItemProperty -Path $eulaRegPath -Name $eulaValue -Value $eulaAccepted -PropertyType String -Force
}
}
}
return $eulaAccepted
}

if ($accepteula)
    {
         ShowEULAIfNeeded "DS Authentication Scripts:" 2
         "EULA Accepted"
    }
else
    {
        $eulaAccepted = ShowEULAIfNeeded "DS Authentication Scripts:" 0
        if($eulaAccepted -ne "Yes")
            {
                "EULA Declined"
                exit
            }
         "EULA Accepted"
    }
###############################################################################
write-host "  _           __  __ ___ ___   ____  __ ___  " -ForegroundColor Yellow
write-host " | |__ _  _  |  \/  / __| _ ) |__ / / /| __| " -ForegroundColor Yellow
write-host " | '_ \ || | | |\/| \__ \ _ \  |_ \/ _ \__ \ " -ForegroundColor Yellow
write-host " |_.__/\_, | |_|  |_|___/___/ |___/\___/___/ " -ForegroundColor Yellow
write-host "       |__/                                  " -ForegroundColor Yellow
Start-Sleep -s 2
write-host ""                                                                                   
write-host ""
write-host ""
write-host ""
###############################################################################


#----------------------------------------------------------------------------------------

param(
    [Parameter(Mandatory = $false)]
    [switch]$WhatIf
)

# Add Windows Forms assembly for file dialog
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Function to show file picker dialog
function Show-FilePickerDialog {
    $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $fileDialog.Title = "Select CSV File with Mailbox Permissions"
    $fileDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    $fileDialog.FilterIndex = 1
    $fileDialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")
    $fileDialog.Multiselect = $false
    
    # Show the dialog
    $result = $fileDialog.ShowDialog()
    
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return $fileDialog.FileName
    }
    else {
        Write-Host "No file selected. Exiting script." -ForegroundColor Yellow
        exit 0
    }
}

# Function to show confirmation dialog
function Show-ConfirmationDialog {
    param(
        [string]$Message,
        [string]$Title = "Confirmation"
    )
    
    $result = [System.Windows.Forms.MessageBox]::Show(
        $Message,
        $Title,
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    
    return $result -eq [System.Windows.Forms.DialogResult]::Yes
}

# Function to show information dialog
function Show-InfoDialog {
    param(
        [string]$Message,
        [string]$Title = "Information"
    )
    
    [System.Windows.Forms.MessageBox]::Show(
        $Message,
        $Title,
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    )
}

# Function to show error dialog
function Show-ErrorDialog {
    param(
        [string]$Message,
        [string]$Title = "Error"
    )
    
    [System.Windows.Forms.MessageBox]::Show(
        $Message,
        $Title,
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )
}

# Show welcome message
$welcomeMessage = @"
Exchange Online Mailbox Permissions Manager

This script will:
• Set Send As and Full Access permissions based on your CSV file
• Remove permissions not defined in the CSV file
• Preserve system permissions

Click OK to select your CSV file.
"@

Show-InfoDialog -Message $welcomeMessage -Title "Welcome"

# Show file picker dialog
Write-Host "Opening file picker dialog..." -ForegroundColor Cyan
$CsvPath = Show-FilePickerDialog

Write-Host "Selected file: $CsvPath" -ForegroundColor Green

# Check if Exchange Online module is available
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    $errorMsg = "ExchangeOnlineManagement module is not installed.`n`nPlease install it using:`nInstall-Module -Name ExchangeOnlineManagement"
    Show-ErrorDialog -Message $errorMsg
    Write-Error "ExchangeOnlineManagement module is not installed. Please install it using: Install-Module -Name ExchangeOnlineManagement"
    exit 1
}

# Import Exchange Online module
Import-Module ExchangeOnlineManagement

# Check if connected to Exchange Online
try {
    $null = Get-OrganizationConfig -ErrorAction Stop
    Write-Host "✓ Connected to Exchange Online" -ForegroundColor Green
}
catch {
    Write-Host "Not connected to Exchange Online. Attempting to connect..." -ForegroundColor Yellow
    
    $connectMsg = "Not connected to Exchange Online.`n`nClick Yes to connect now, or No to exit."
    if (Show-ConfirmationDialog -Message $connectMsg -Title "Exchange Online Connection") {
        try {
            Connect-ExchangeOnline -ShowProgress $true
            Write-Host "✓ Successfully connected to Exchange Online" -ForegroundColor Green
        }
        catch {
            $errorMsg = "Failed to connect to Exchange Online:`n$($_.Exception.Message)"
            Show-ErrorDialog -Message $errorMsg
            Write-Error "Failed to connect to Exchange Online: $($_.Exception.Message)"
            exit 1
        }
    }
    else {
        Write-Host "User chose not to connect. Exiting." -ForegroundColor Yellow
        exit 0
    }
}

# Validate CSV file exists
if (-not (Test-Path $CsvPath)) {
    $errorMsg = "CSV file not found:`n$CsvPath"
    Show-ErrorDialog -Message $errorMsg
    Write-Error "CSV file not found: $CsvPath"
    exit 1
}

# Import CSV data
try {
    $csvData = Import-Csv -Path $CsvPath
    Write-Host "✓ Successfully imported CSV file with $($csvData.Count) rows" -ForegroundColor Green
}
catch {
    $errorMsg = "Failed to import CSV file:`n$($_.Exception.Message)"
    Show-ErrorDialog -Message $errorMsg
    Write-Error "Failed to import CSV file: $($_.Exception.Message)"
    exit 1
}

# Validate CSV headers
$requiredHeaders = @('MailboxIdentity', 'UserIdentity', 'SendAs', 'FullAccess')
$csvHeaders = $csvData[0].PSObject.Properties.Name

$missingHeaders = @()
foreach ($header in $requiredHeaders) {
    if ($header -notin $csvHeaders) {
        $missingHeaders += $header
    }
}

if ($missingHeaders.Count -gt 0) {
    $errorMsg = "Missing required columns in CSV:`n$($missingHeaders -join ', ')`n`nRequired columns:`n$($requiredHeaders -join ', ')"
    Show-ErrorDialog -Message $errorMsg
    Write-Error "Missing required columns in CSV: $($missingHeaders -join ', ')"
    Write-Host "Required columns: $($requiredHeaders -join ', ')" -ForegroundColor Yellow
    exit 1
}

# Show summary and confirmation
$mailboxGroups = $csvData | Group-Object -Property MailboxIdentity
$totalPermissions = $csvData.Count

$summaryMessage = @"
CSV File Summary:
• File: $(Split-Path $CsvPath -Leaf)
• Total mailboxes: $($mailboxGroups.Count)
• Total permission entries: $totalPermissions
• Mode: $(if ($WhatIf) { "Preview Mode (No changes will be made)" } else { "Live Mode (Changes will be applied)" })

Do you want to proceed?
"@

if (-not (Show-ConfirmationDialog -Message $summaryMessage -Title "Confirm Processing")) {
    Write-Host "User cancelled operation." -ForegroundColor Yellow
    exit 0
}

Write-Host "`nProcessing $($mailboxGroups.Count) mailboxes..." -ForegroundColor Cyan

$processedMailboxes = 0
$successfulChanges = 0
$errors = @()

foreach ($mailboxGroup in $mailboxGroups) {
    $mailboxIdentity = $mailboxGroup.Name
    $permissions = $mailboxGroup.Group
    
    Write-Host "`n--- Processing mailbox: $mailboxIdentity ---" -ForegroundColor Yellow
    $processedMailboxes++
    
    # Verify mailbox exists
    try {
        $mailbox = Get-Mailbox -Identity $mailboxIdentity -ErrorAction Stop
        Write-Host "✓ Mailbox found: $($mailbox.DisplayName)" -ForegroundColor Green
    }
    catch {
        $errorMsg = "Mailbox not found: $mailboxIdentity"
        Write-Warning $errorMsg
        $errors += $errorMsg
        continue
    }
    
    # Get current permissions
    try {
        $currentFullAccess = Get-MailboxPermission -Identity $mailboxIdentity | 
            Where-Object { $_.User -notlike "NT AUTHORITY\SELF" -and $_.User -notlike "S-1-*" -and $_.AccessRights -contains "FullAccess" }
        
        $currentSendAs = Get-RecipientPermission -Identity $mailboxIdentity | 
            Where-Object { $_.Trustee -notlike "NT AUTHORITY\SELF" -and $_.Trustee -notlike "S-1-*" -and $_.AccessRights -contains "SendAs" }
    }
    catch {
        $errorMsg = "Failed to get current permissions for $mailboxIdentity : $($_.Exception.Message)"
        Write-Warning $errorMsg
        $errors += $errorMsg
        continue
    }
    
    # Process desired permissions from CSV
    $desiredFullAccess = @()
    $desiredSendAs = @()
    
    foreach ($permission in $permissions) {
        if ($permission.FullAccess -eq "TRUE" -or $permission.FullAccess -eq "Yes" -or $permission.FullAccess -eq "1") {
            $desiredFullAccess += $permission.UserIdentity
        }
        if ($permission.SendAs -eq "TRUE" -or $permission.SendAs -eq "Yes" -or $permission.SendAs -eq "1") {
            $desiredSendAs += $permission.UserIdentity
        }
    }
    
    # Process Full Access permissions
    Write-Host "Processing Full Access permissions..." -ForegroundColor Cyan
    
    # Add missing Full Access permissions
    foreach ($user in $desiredFullAccess) {
        if ($currentFullAccess.User -notcontains $user) {
            Write-Host "  Adding Full Access for: $user" -ForegroundColor Green
            if (-not $WhatIf) {
                try {
                    Add-MailboxPermission -Identity $mailboxIdentity -User $user -AccessRights FullAccess -InheritanceType All -Confirm:$false
                    Write-Host "  ✓ Successfully added Full Access for $user" -ForegroundColor Green
                    $successfulChanges++
                }
                catch {
                    $errorMsg = "Failed to add Full Access for $user : $($_.Exception.Message)"
                    Write-Warning "  $errorMsg"
                    $errors += $errorMsg
                }
            }
            else {
                Write-Host "  [WHATIF] Would add Full Access for: $user" -ForegroundColor Magenta
            }
        }
        else {
            Write-Host "  Full Access already exists for: $user" -ForegroundColor Gray
        }
    }
    
    # Remove unwanted Full Access permissions
    foreach ($currentPerm in $currentFullAccess) {
        if ($currentPerm.User -notin $desiredFullAccess) {
            Write-Host "  Removing Full Access for: $($currentPerm.User)" -ForegroundColor Red
            if (-not $WhatIf) {
                try {
                    Remove-MailboxPermission -Identity $mailboxIdentity -User $currentPerm.User -AccessRights FullAccess -Confirm:$false
                    Write-Host "  ✓ Successfully removed Full Access for $($currentPerm.User)" -ForegroundColor Green
                    $successfulChanges++
                }
                catch {
                    $errorMsg = "Failed to remove Full Access for $($currentPerm.User) : $($_.Exception.Message)"
                    Write-Warning "  $errorMsg"
                    $errors += $errorMsg
                }
            }
            else {
                Write-Host "  [WHATIF] Would remove Full Access for: $($currentPerm.User)" -ForegroundColor Magenta
            }
        }
    }
    
    # Process Send As permissions
    Write-Host "Processing Send As permissions..." -ForegroundColor Cyan
    
    # Add missing Send As permissions
    foreach ($user in $desiredSendAs) {
        if ($currentSendAs.Trustee -notcontains $user) {
            Write-Host "  Adding Send As for: $user" -ForegroundColor Green
            if (-not $WhatIf) {
                try {
                    Add-RecipientPermission -Identity $mailboxIdentity -Trustee $user -AccessRights SendAs -Confirm:$false
                    Write-Host "  ✓ Successfully added Send As for $user" -ForegroundColor Green
                    $successfulChanges++
                }
                catch {
                    $errorMsg = "Failed to add Send As for $user : $($_.Exception.Message)"
                    Write-Warning "  $errorMsg"
                    $errors += $errorMsg
                }
            }
            else {
                Write-Host "  [WHATIF] Would add Send As for: $user" -ForegroundColor Magenta
            }
        }
        else {
            Write-Host "  Send As already exists for: $user" -ForegroundColor Gray
        }
    }
    
    # Remove unwanted Send As permissions
    foreach ($currentPerm in $currentSendAs) {
        if ($currentPerm.Trustee -notin $desiredSendAs) {
            Write-Host "  Removing Send As for: $($currentPerm.Trustee)" -ForegroundColor Red
            if (-not $WhatIf) {
                try {
                    Remove-RecipientPermission -Identity $currentPerm.Identity -Trustee $currentPerm.Trustee -AccessRights SendAs -Confirm:$false
                    Write-Host "  ✓ Successfully removed Send As for $($currentPerm.Trustee)" -ForegroundColor Green
                    $successfulChanges++
                }
                catch {
                    $errorMsg = "Failed to remove Send As for $($currentPerm.Trustee) : $($_.Exception.Message)"
                    Write-Warning "  $errorMsg"
                    $errors += $errorMsg
                }
            }
            else {
                Write-Host "  [WHATIF] Would remove Send As for: $($currentPerm.Trustee)" -ForegroundColor Magenta
            }
        }
    }
}

Write-Host "`n✓ Script execution completed!" -ForegroundColor Green

# Show completion summary
$completionMessage = @"
Processing Complete!

Summary:
• Processed mailboxes: $processedMailboxes
• Successful changes: $successfulChanges
• Errors encountered: $($errors.Count)
$(if ($WhatIf) { "`nNote: This was a preview run. No actual changes were made." } else { "" })
$(if ($errors.Count -gt 0) { "`nErrors:`n" + ($errors -join "`n") } else { "" })
"@

if ($errors.Count -gt 0) {
    Show-ErrorDialog -Message $completionMessage -Title "Processing Complete (With Errors)"
}
else {
    Show-InfoDialog -Message $completionMessage -Title "Processing Complete"
}

if ($WhatIf) {
    Write-Host "`nNote: This was a WhatIf run. No changes were made." -ForegroundColor Yellow
    Write-Host "Remove the -WhatIf parameter to apply the changes." -ForegroundColor Yellow
}

Write-Host "`nPress any key to exit..." -ForegroundColor Gray
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
