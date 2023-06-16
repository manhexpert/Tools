# Load required assemblies
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

# Variables
$RBCM_log_Path      = "$env:ProgramData\Bosch\RBcm\Logs\packages\"
$RBCM_cfg_Path      = "$env:ProgramData\Bosch\RBcm\config\"
$GUIDx64            = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"
$GUIDX86            = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\"
$Source_PackagePath = "\\rb-sccm-pkg-d.bosch.com\rbcm$\"
$FingePrint_Path    = "HKLM:\SOFTWARE\BOSCH\RBcm\packages\"
$PMI                = "HKLM:\SOFTWARE\BOSCH\RBcm\PMI\Client\Environment"
$Daten              = "$env:SystemDrive\daten\"
$Active_Setup_HKLM  = "HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\"
$Active_Setup_HKCU  = "HKCU:\SOFTWARE\Microsoft\Active Setup\Installed Components"
$Registry_Firewall  = "HKLM:\SYSTEM\CurrentControlSet\Services\SharedAccess\Parameters\FirewallPolicy\FirewallRules"
$Archive            = "\\si0vm1384.de.bosch.com\archive$\"

# Create form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Tool Test package"
$form.AutoSize = $true
$form.StartPosition = "CenterScreen"
$form.MaximizeBox = $false
$form.FormBorderStyle = "FixedSingle"
$form.Padding = New-Object System.Windows.Forms.Padding(15, 15, 15, 15)


# Create label
$rbcmlabel = New-Object System.Windows.Forms.Label
$rbcmlabel.Location = New-Object System.Drawing.Point(20, 20)
$rbcmlabel.Size = New-Object System.Drawing.Size(100, 20)
$rbcmlabel.Text = "RBCM Number:"
$form.Controls.Add($rbcmlabel)

# Create text box
$textrbcm = New-Object System.Windows.Forms.TextBox
$textrbcm.Location = New-Object System.Drawing.Point(150, 20)
$textrbcm.Size = New-Object System.Drawing.Size(100, 20)
$form.Controls.Add($textrbcm)
$textrbcm.Add_TextChanged({
    $global:RBCM = $textrbcm.Text
    $global:RBCM_log = "$RBCM_log_Path`RBCM$RBCM"
    $global:RBCM_cfg = "$RBCM_cfg_Path`RBCM$RBCM"
    $global:Source_Package = "$Source_PackagePath`RBCM$RBCM"
    $global:FingePrint = "$FingePrint_Path`RBCM$RBCM"
})

#function
Function OpenRegistry ([string]$ResPath)
{
    $regPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Applets\Regedit"
    $name = "LastKey"
    $value = "Computer\"+(Convert-Path ($ResPath))

    New-ItemProperty -Path $regPath -Name $name -Value $value -PropertyType String -Force | Out-Null
    Start-Process RegEdit
}

Function Create-Button {
    param(
        [System.Drawing.Point]$Location,
        [System.Drawing.Size]$Size,
        [string]$Text,
        [ScriptBlock]$Action
    )

    $button = New-Object System.Windows.Forms.Button
    $button.Location = $Location
    $button.Size = $Size
    $button.Text = $Text
    $button.Add_Click($Action)

    return $button
}

Function RunTaskscheduler ([string]$taskName,[string]$taskPath)
{
	Start-Process -FilePath "schtasks.exe" -ArgumentList "/run /tn ""$taskPath\$taskName""" -Verb RunAs	-WindowStyle Hidden
}

Function CopyData([string]$From,[string]$To)
{
# Create check box and scroll
$sourceFolder = $From
$destinationFolder = $To

$subFolders = Get-ChildItem -Path $sourceFolder -Directory

# Create checkbox to select folder
$scrollBox = New-Object System.Windows.Forms.Panel
$scrollBox.Width = 270
$scrollBox.Height = 280
$scrollBox.AutoScroll = $true

$x = 10
$y = 10
$checkBoxes = @()

foreach ($subFolder in $subFolders) {
    $checkBox = New-Object System.Windows.Forms.CheckBox
    $checkBox.Text = $subFolder.Name
    $checkBox.AutoSize = $true
    $checkBox.Location = New-Object System.Drawing.Point($x, $y)
    $scrollBox.Controls.Add($checkBox)
    $checkBoxes += $checkBox
    $y += 25 # 25 pixels apart
}

# Create dialog box
$form = New-Object System.Windows.Forms.Form
$form.Text = "Select Folders to Copy"
$form.Width = 300
$form.Height = 400
$Form.StartPosition = "CenterParent"

$okButton = New-Object System.Windows.Forms.Button
$okButton.Text = "OK"
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$okButton.Location = New-Object System.Drawing.Point(60, 300)
$okButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$form.AcceptButton = $okButton
$form.Controls.Add($okButton)

$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Text = "Cancel"
$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$cancelButton.Location = New-Object System.Drawing.Point(150,300)
$cancelButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
$form.CancelButton = $cancelButton
$form.Controls.Add($cancelButton)

# Add scroll bar to dialog
$scrollBox.Location = New-Object System.Drawing.Point(10, 10)
$form.Controls.Add($scrollBox)

$result = $form.ShowDialog()

# Copy selected folders if user selects OK
if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
    foreach ($checkBox in $checkBoxes) {
        if ($checkBox.Checked) 
        {
            $subFolderName = $checkBox.Text
            $sourceSubFolder = "$sourceFolder\$subFolderName"
            $objShell = New-Object -ComObject "Shell.Application"
            $objFolder = $objShell.NameSpace($destinationFolder) 
            $objFolder.CopyHere($sourceSubFolder,16)
        }
    }
    explorer "$destinationFolder"
}

}
Function OpenFolder([string]$Path)
{
    if (Test-Path -Path "$Path" -PathType Any)
    {
        explorer "$Path"
    }
    Else
    {
    [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    [System.Windows.Forms.MessageBox]::Show("Folder not found", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }

}


# Create button colum 1 and bind to event handler
$buttons = @(
    @{ Name = "RBCM LOG"; Value = { OpenFolder -Path "$RBCM_Log" } },
    @{ Name = "RBCM Config"; Value = { OpenFolder -Path "$RBCM_Cfg" } },
    @{ Name = "GUIDx64"; Value = { OpenRegistry -ResPath $GUIDx64 } },
    @{ Name = "GUIDx86"; Value = { OpenRegistry -ResPath $GUIDx86 } },
    @{ Name = "Source Package"; Value = { OpenFolder -Path "$Source_Package" } },
    @{ Name = "FingePrint"; Value = { OpenRegistry -ResPath $FingePrint } },
    @{ Name = "PMI"; Value = { OpenRegistry -ResPath $PMI } },
    @{ Name = "Copy to daten"; Value = { CopyData -From $Source_Package -To $Daten} },
    @{ Name = "Archive"; Value = { OpenFolder -Path "$Archive" } }
    
)
$i=0
foreach ($button in $buttons) 
{
    $Vertical = 80 + $i * 40
    $location = New-Object System.Drawing.Point(20,$Vertical )
    $size = New-Object System.Drawing.Size(100, 30)
    $name = $button.Name
    $Action = $button.Value
    $button = Create-Button -Location $location -Size $size -Text $name -Action $Action
    $form.Controls.Add($button)
    $i++
}

# Create button colum 2 and bind to event handler
$buttons = @(
    @{ Name = "Active HKLM"; Value = { OpenRegistry -ResPath $Active_Setup_HKLM } },
    @{ Name = "Active HKCU"; Value = { OpenRegistry -ResPath $Active_Setup_HKCU } },
    @{ Name = "Fire wall"; Value = { OpenRegistry -ResPath $Registry_Firewall } },
    @{ Name = "Env variable"; Value = { Start-Process rundll32 -ArgumentList "sysdm.cpl,EditEnvironmentVariables"} },
    @{ Name = "Smscfgrc"; Value = { cmd /c "control smscfgrc"} },
    @{ Name = "Task scheduler"; Value = { Start-Process "$env:windir\System32\taskschd.msc" -Verb RunAs } }, 
    @{ Name = "Get PMI"; Value = { RunTaskscheduler -taskName "PmiGetConfig" -taskPath "\RBCM" } },
    @{ Name = "Run RBCM0048"; Value = { RunTaskscheduler -taskName "install_always" -taskPath "\RBCM\RBCM0048"} },
    @{ Name = "Control panel"; Value = { control appwiz.cpl} },
    @{ Name = "Security policy"; Value = { Start-Process "C:\Windows\System32\secpol.msc" -Verb RunAs} }
)
$i=0
foreach ($button in $buttons) 
{
    $Vertical = 80 + $i * 40
    $location = New-Object System.Drawing.Point(150,$Vertical )
    $size = New-Object System.Drawing.Size(100, 30)
    $name = $button.Name
    $Action = $button.Value
    $button = Create-Button -Location $location -Size $size -Text $name -Action $Action
    $form.Controls.Add($button)
    $i++
}

# Show form
$form.ShowDialog() | Out-Null
Start-Process "C:\ProgramData\Bosch\RBcm\PMI\PMIGetConfig.exe" -ArgumentList "/v"
