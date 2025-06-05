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
    Manages Exchange Online mailbox permissions based on CSV file input with GUI file picker and HTML reporting
.DESCRIPTION
    This script sets "Send As" and "Full Access" permissions for mailboxes defined in a CSV file.
    It removes any existing permissions that are not defined in the CSV file.
    Features a GUI file picker for easy CSV file selection and generates detailed HTML reports.
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

# Global variables for reporting
$global:reportData = @{
    StartTime = Get-Date
    EndTime = $null
    CsvFile = ""
    TotalMailboxes = 0
    ProcessedMailboxes = 0
    SuccessfulChanges = 0
    Errors = @()
    MailboxResults = @()
    WhatIfMode = $false
}

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

# Function to add mailbox result to report data
function Add-MailboxResult {
    param(
        [string]$MailboxIdentity,
        [string]$MailboxDisplayName,
        [array]$PermissionChanges,
        [array]$Errors,
        [string]$Status
    )
    
    $mailboxResult = @{
        MailboxIdentity = $MailboxIdentity
        MailboxDisplayName = $MailboxDisplayName
        PermissionChanges = $PermissionChanges
        Errors = $Errors
        Status = $Status
        ProcessedAt = Get-Date
    }
    
    $global:reportData.MailboxResults += $mailboxResult
}

# Function to add permission change to tracking
function Add-PermissionChange {
    param(
        [string]$Action,
        [string]$PermissionType,
        [string]$User,
        [string]$Status,
        [string]$ErrorMessage = ""
    )
    
    return @{
        Action = $Action
        PermissionType = $PermissionType
        User = $User
        Status = $Status
        ErrorMessage = $ErrorMessage
        Timestamp = Get-Date
    }
}

# Function to generate HTML report
function Generate-HTMLReport {
    param(
        [string]$OutputPath
    )
    
    $global:reportData.EndTime = Get-Date
    $duration = $global:reportData.EndTime - $global:reportData.StartTime
    
    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Exchange Online Permissions Report</title>
    <style>
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            margin: 0;
            padding: 20px;
            background-color: #f5f5f5;
            color: #333;
        }
        .container {
            max-width: 1200px;
            margin: 0 auto;
            background-color: white;
            border-radius: 8px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            overflow: hidden;
        }
        .header {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 30px;
            text-align: center;
        }
        .header h1 {
            margin: 0;
            font-size: 2.5em;
            font-weight: 300;
        }
        .header p {
            margin: 10px 0 0 0;
            opacity: 0.9;
            font-size: 1.1em;
        }
        .summary {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            padding: 30px;
            background-color: #f8f9fa;
        }
        .summary-card {
            background: white;
            padding: 20px;
            border-radius: 8px;
            text-align: center;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            border-left: 4px solid #667eea;
        }
        .summary-card h3 {
            margin: 0 0 10px 0;
            color: #667eea;
            font-size: 2em;
        }
        .summary-card p {
            margin: 0;
            color: #666;
            font-weight: 500;
        }
        .content {
            padding: 30px;
        }
        .mailbox-section {
            margin-bottom: 30px;
            border: 1px solid #e0e0e0;
            border-radius: 8px;
            overflow: hidden;
        }
        .mailbox-header {
            background-color: #f8f9fa;
            padding: 15px 20px;
            border-bottom: 1px solid #e0e0e0;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        .mailbox-title {
            font-weight: 600;
            font-size: 1.1em;
            color: #333;
        }
        .status-badge {
            padding: 4px 12px;
            border-radius: 20px;
            font-size: 0.85em;
            font-weight: 500;
            text-transform: uppercase;
        }
        .status-success {
            background-color: #d4edda;
            color: #155724;
            border: 1px solid #c3e6cb;
        }
        .status-error {
            background-color: #f8d7da;
            color: #721c24;
            border: 1px solid #f5c6cb;
        }
        .status-warning {
            background-color: #fff3cd;
            color: #856404;
            border: 1px solid #ffeaa7;
        }
        .permissions-table {
            width: 100%;
            border-collapse: collapse;
            margin: 0;
        }
        .permissions-table th {
            background-color: #f8f9fa;
            padding: 12px;
            text-align: left;
            font-weight: 600;
            color: #495057;
            border-bottom: 2px solid #dee2e6;
        }
        .permissions-table td {
            padding: 12px;
            border-bottom: 1px solid #dee2e6;
            vertical-align: top;
        }
        .permissions-table tr:hover {
            background-color: #f8f9fa;
        }
        .action-added {
            color: #28a745;
            font-weight: 500;
        }
        .action-removed {
            color: #dc3545;
            font-weight: 500;
        }
        .action-unchanged {
            color: #6c757d;
            font-style: italic;
        }
        .permission-type {
            display: inline-block;
            padding: 2px 8px;
            border-radius: 4px;
            font-size: 0.85em;
            font-weight: 500;
        }
        .permission-fullaccess {
            background-color: #e3f2fd;
            color: #1565c0;
        }
        .permission-sendas {
            background-color: #f3e5f5;
            color: #7b1fa2;
        }
        .error-section {
            background-color: #f8d7da;
            border: 1px solid #f5c6cb;
            border-radius: 8px;
            padding: 15px;
            margin-top: 20px;
        }
        .error-title {
            color: #721c24;
            font-weight: 600;
            margin-bottom: 10px;
        }
        .error-list {
            list-style: none;
            padding: 0;
            margin: 0;
        }
        .error-list li {
            padding: 5px 0;
            color: #721c24;
        }
        .whatif-notice {
            background-color: #fff3cd;
            border: 1px solid #ffeaa7;
            border-radius: 8px;
            padding: 15px;
            margin-bottom: 20px;
            color: #856404;
            font-weight: 500;
            text-align: center;
        }
        .footer {
            background-color: #f8f9fa;
            padding: 20px;
            text-align: center;
            color: #666;
            border-top: 1px solid #e0e0e0;
        }
        .no-changes {
            text-align: center;
            padding: 40px;
            color: #666;
            font-style: italic;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>Exchange Online Permissions Report</h1>
            <p>Generated on $($global:reportData.EndTime.ToString('dddd, MMMM dd, yyyy at HH:mm:ss'))</p>
        </div>
        
        $(if ($global:reportData.WhatIfMode) {
            '<div class="whatif-notice">
                <strong>PREVIEW MODE:</strong> This report shows what would have been changed. No actual modifications were made.
            </div>'
        })
        
        <div class="summary">
            <div class="summary-card">
                <h3>$($global:reportData.TotalMailboxes)</h3>
                <p>Total Mailboxes</p>
            </div>
            <div class="summary-card">
                <h3>$($global:reportData.ProcessedMailboxes)</h3>
                <p>Processed Successfully</p>
            </div>
            <div class="summary-card">
                <h3>$($global:reportData.SuccessfulChanges)</h3>
                <p>Permission Changes</p>
            </div>
            <div class="summary-card">
                <h3>$($global:reportData.Errors.Count)</h3>
                <p>Errors Encountered</p>
            </div>
        </div>
        
        <div class="content">
            <h2>Execution Details</h2>
            <table class="permissions-table">
                <tr>
                    <td><strong>CSV File:</strong></td>
                    <td>$(Split-Path $global:reportData.CsvFile -Leaf)</td>
                </tr>
                <tr>
                    <td><strong>Start Time:</strong></td>
                    <td>$($global:reportData.StartTime.ToString('yyyy-MM-dd HH:mm:ss'))</td>
                </tr>
                <tr>
                    <td><strong>End Time:</strong></td>
                    <td>$($global:reportData.EndTime.ToString('yyyy-MM-dd HH:mm:ss'))</td>
                </tr>
                <tr>
                    <td><strong>Duration:</strong></td>
                    <td>$($duration.ToString('mm\:ss'))</td>
                </tr>
                <tr>
                    <td><strong>Mode:</strong></td>
                    <td>$(if ($global:reportData.WhatIfMode) { "Preview Mode (WhatIf)" } else { "Live Mode" })</td>
                </tr>
            </table>
            
            <h2>Mailbox Processing Results</h2>
            
            $(if ($global:reportData.MailboxResults.Count -eq 0) {
                '<div class="no-changes">No mailboxes were processed.</div>'
            } else {
                $global:reportData.MailboxResults | ForEach-Object {
                    $mailbox = $_
                    $statusClass = switch ($mailbox.Status) {
                        "Success" { "status-success" }
                        "Error" { "status-error" }
                        "Warning" { "status-warning" }
                        default { "status-warning" }
                    }
                    
                    "<div class='mailbox-section'>
                        <div class='mailbox-header'>
                            <div class='mailbox-title'>$($mailbox.MailboxDisplayName) ($($mailbox.MailboxIdentity))</div>
                            <div class='status-badge $statusClass'>$($mailbox.Status)</div>
                        </div>"
                    
                    if ($mailbox.PermissionChanges.Count -gt 0) {
                        "<table class='permissions-table'>
                            <thead>
                                <tr>
                                    <th>Action</th>
                                    <th>Permission Type</th>
                                    <th>User</th>
                                    <th>Status</th>
                                    <th>Details</th>
                                </tr>
                            </thead>
                            <tbody>"
                        
                        $mailbox.PermissionChanges | ForEach-Object {
                            $change = $_
                            $actionClass = switch ($change.Action) {
                                "Added" { "action-added" }
                                "Removed" { "action-removed" }
                                default { "action-unchanged" }
                            }
                            $permissionClass = switch ($change.PermissionType) {
                                "Full Access" { "permission-fullaccess" }
                                "Send As" { "permission-sendas" }
                                default { "" }
                            }
                            
                            "<tr>
                                <td class='$actionClass'>$($change.Action)</td>
                                <td><span class='permission-type $permissionClass'>$($change.PermissionType)</span></td>
                                <td>$($change.User)</td>
                                <td class='$(if ($change.Status -eq "Success") { "action-added" } else { "action-removed" })'>$($change.Status)</td>
                                <td>$(if ($change.ErrorMessage) { $change.ErrorMessage } else { "Completed successfully" })</td>
                            </tr>"
                        }
                        
                        "</tbody></table>"
                    } else {
                        "<div class='no-changes'>No permission changes were needed for this mailbox.</div>"
                    }
                    
                    if ($mailbox.Errors.Count -gt 0) {
                        "<div class='error-section'>
                            <div class='error-title'>Errors for this mailbox:</div>
                            <ul class='error-list'>"
                        $mailbox.Errors | ForEach-Object {
                            "<li>$_</li>"
                        }
                        "</ul></div>"
                    }
                    
                    "</div>"
                }
            })
            
            $(if ($global:reportData.Errors.Count -gt 0) {
                "<h2>Global Errors</h2>
                <div class='error-section'>
                    <div class='error-title'>The following errors occurred during processing:</div>
                    <ul class='error-list'>"
                $global:reportData.Errors | ForEach-Object {
                    "<li>$_</li>"
                }
                "</ul></div>"
            })
        </div>
        
        <div class="footer">
            <p>Report generated by Exchange Online Permissions Manager</p>
            <p>PowerShell Script executed by $($env:USERNAME) on $($env:COMPUTERNAME)</p>
        </div>
    </div>
</body>
</html>
"@

    try {
        $html | Out-File -FilePath $OutputPath -Encoding UTF8
        return $true
    }
    catch {
        Write-Error "Failed to generate HTML report: $($_.Exception.Message)"
        return $false
    }
}

# Show welcome message
$welcomeMessage = @"
Exchange Online Mailbox Permissions Manager

This script will:
• Set Send As and Full Access permissions based on your CSV file
• Remove permissions not defined in the CSV file
• Preserve system permissions
• Generate a detailed HTML report

Click OK to select your CSV file.
"@

Show-InfoDialog -Message $welcomeMessage -Title "Welcome"

# Show file picker dialog
Write-Host "Opening file picker dialog..." -ForegroundColor Cyan
$CsvPath = Show-FilePickerDialog
$global:reportData.CsvFile = $CsvPath
$global:reportData.WhatIfMode = $WhatIf

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
$global:reportData.TotalMailboxes = $mailboxGroups.Count

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

foreach ($mailboxGroup in $mailboxGroups) {
    $mailboxIdentity = $mailboxGroup.Name
    $permissions = $mailboxGroup.Group
    $mailboxErrors = @()
    $permissionChanges = @()
    
    Write-Host "`n--- Processing mailbox: $mailboxIdentity ---" -ForegroundColor Yellow
    
    # Verify mailbox exists
    try {
        $mailbox = Get-Mailbox -Identity $mailboxIdentity -ErrorAction Stop
        Write-Host "✓ Mailbox found: $($mailbox.DisplayName)" -ForegroundColor Green
        $mailboxDisplayName = $mailbox.DisplayName
    }
    catch {
        $errorMsg = "Mailbox not found: $mailboxIdentity"
        Write-Warning $errorMsg
        $global:reportData.Errors += $errorMsg
        $mailboxErrors += $errorMsg
        Add-MailboxResult -MailboxIdentity $mailboxIdentity -MailboxDisplayName $mailboxIdentity -PermissionChanges @() -Errors $mailboxErrors -Status "Error"
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
        $global:reportData.Errors += $errorMsg
        $mailboxErrors += $errorMsg
        Add-MailboxResult -MailboxIdentity $mailboxIdentity -MailboxDisplayName $mailboxDisplayName -PermissionChanges @() -Errors $mailboxErrors -Status "Error"
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
                    $global:reportData.SuccessfulChanges++
                    $permissionChanges += Add-PermissionChange -Action "Added" -PermissionType "Full Access" -User $user -Status "Success"
                }
                catch {
                    $errorMsg = "Failed to add Full Access for $user : $($_.Exception.Message)"
                    Write-Warning "  $errorMsg"
                    $mailboxErrors += $errorMsg
                    $permissionChanges += Add-PermissionChange -Action "Added" -PermissionType "Full Access" -User $user -Status "Failed" -ErrorMessage $_.Exception.Message
                }
            }
            else {
                Write-Host "  [WHATIF] Would add Full Access for: $user" -ForegroundColor Magenta
                $permissionChanges += Add-PermissionChange -Action "Added" -PermissionType "Full Access" -User $user -Status "Preview"
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
                    $global:reportData.SuccessfulChanges++
                    $permissionChanges += Add-PermissionChange -Action "Removed" -PermissionType "Full Access" -User $currentPerm.User -Status "Success"
                }
                catch {
                    $errorMsg = "Failed to remove Full Access for $($currentPerm.User) : $($_.Exception.Message)"
                    Write-Warning "  $errorMsg"
                    $mailboxErrors += $errorMsg
                    $permissionChanges += Add-PermissionChange -Action "Removed" -PermissionType "Full Access" -User $currentPerm.User -Status "Failed" -ErrorMessage $_.Exception.Message
                }
            }
            else {
                Write-Host "  [WHATIF] Would remove Full Access for: $($currentPerm.User)" -ForegroundColor Magenta
                $permissionChanges += Add-PermissionChange -Action "Removed" -PermissionType "Full Access" -User $currentPerm.User -Status "Preview"
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
                    $global:reportData.SuccessfulChanges++
                    $permissionChanges += Add-PermissionChange -Action "Added" -PermissionType "Send As" -User $user -Status "Success"
                }
                catch {
                    $errorMsg = "Failed to add Send As for $user : $($_.Exception.Message)"
                    Write-Warning "  $errorMsg"
                    $mailboxErrors += $errorMsg
                    $permissionChanges += Add-PermissionChange -Action "Added" -PermissionType "Send As" -User $user -Status "Failed" -ErrorMessage $_.Exception.Message
                }
            }
            else {
                Write-Host "  [WHATIF] Would add Send As for: $user" -ForegroundColor Magenta
                $permissionChanges += Add-PermissionChange -Action "Added" -PermissionType "Send As" -User $user -Status "Preview"
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
                    $global:reportData.SuccessfulChanges++
                    $permissionChanges += Add-PermissionChange -Action "Removed" -PermissionType "Send As" -User $currentPerm.Trustee -Status "Success"
                }
                catch {
                    $errorMsg = "Failed to remove Send As for $($currentPerm.Trustee) : $($_.Exception.Message)"
                    Write-Warning "  $errorMsg"
                    $mailboxErrors += $errorMsg
                    $permissionChanges += Add-PermissionChange -Action "Removed" -PermissionType "Send As" -User $currentPerm.Trustee -Status "Failed" -ErrorMessage $_.Exception.Message
                }
            }
            else {
                Write-Host "  [WHATIF] Would remove Send As for: $($currentPerm.Trustee)" -ForegroundColor Magenta
                $permissionChanges += Add-PermissionChange -Action "Removed" -PermissionType "Send As" -User $currentPerm.Trustee -Status "Preview"
            }
        }
    }
    
    # Determine mailbox status
    $mailboxStatus = if ($mailboxErrors.Count -gt 0) { "Error" } 
                    elseif ($permissionChanges.Count -eq 0) { "No Changes" }
                    else { "Success" }
    
    # Add mailbox result to report
    Add-MailboxResult -MailboxIdentity $mailboxIdentity -MailboxDisplayName $mailboxDisplayName -PermissionChanges $permissionChanges -Errors $mailboxErrors -Status $mailboxStatus
    $global:reportData.ProcessedMailboxes++
}

Write-Host "`n✓ Script execution completed!" -ForegroundColor Green

# Generate HTML report
$csvDirectory = Split-Path -Path $CsvPath -Parent
$csvFileName = [System.IO.Path]::GetFileNameWithoutExtension($CsvPath)
$reportFileName = "$csvFileName-Permissions-Report-$(Get-Date -Format 'yyyyMMdd-HHmmss').html"
$reportPath = Join-Path -Path $csvDirectory -ChildPath $reportFileName

Write-Host "`nGenerating HTML report..." -ForegroundColor Cyan

if (Generate-HTMLReport -OutputPath $reportPath) {
    Write-Host "✓ HTML report generated successfully: $reportPath" -ForegroundColor Green
    
    # Ask user if they want to open the report
    $openReportMsg = "HTML report has been generated successfully!`n`nLocation: $reportPath`n`nWould you like to open the report now?"
    if (Show-ConfirmationDialog -Message $openReportMsg -Title "Report Generated") {
        try {
            Start-Process $reportPath
        }
        catch {
            Write-Warning "Could not open report automatically. Please open it manually from: $reportPath"
        }
    }
}
else {
    Write-Error "Failed to generate HTML report"
}

# Show completion summary
$completionMessage = @"
Processing Complete!

Summary:
• Processed mailboxes: $($global:reportData.ProcessedMailboxes)
• Successful changes: $($global:reportData.SuccessfulChanges)
• Errors encountered: $($global:reportData.Errors.Count)
$(if ($WhatIf) { "`nNote: This was a preview run. No actual changes were made." } else { "" })

HTML Report: $reportFileName
Location: $(Split-Path -Path $CsvPath -Parent)
"@

if ($global:reportData.Errors.Count -gt 0) {
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
