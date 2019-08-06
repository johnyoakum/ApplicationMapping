#From homw on VPN, script ran in 53 minutes from time of kick off until it showed the form.

Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process
$ScriptPathParent = split-path -Parent -Path $MyInvocation.MyCommand.Definition


# *****************Site configuration - Coinfigure the following three lines to match your site*****************
$SiteCode = "049" # Site code 
$ProviderMachineName = "anc-sccm-site01.ua.ad.alaska.edu" # SMS Provider machine name
$QueryName = "All Software Installed at UAA" # Pre-Defined query that gets all installed software listed in SCCM

# Customizations
$initParams = @{}
$AllApplicationMappings = @()

$CSVExists = Test-path -Path "$ScriptPathParent\Applications.csv"
$Debug = $true

function msgbox {
param (
    [string]$Message,
    [string]$Title = 'Message box title',   
    [string]$buttons = 'OKCancel'
)
# This function displays a message box by calling the .Net Windows.Forms (MessageBox class)
 
# Load the assembly
Add-Type -AssemblyName System.Windows.Forms | Out-Null
 
# Define the button types
switch ($buttons) {
   'ok' {$btn = [System.Windows.Forms.MessageBoxButtons]::OK; break}
   'okcancel' {$btn = [System.Windows.Forms.MessageBoxButtons]::OKCancel; break}
   'AbortRetryIgnore' {$btn = [System.Windows.Forms.MessageBoxButtons]::AbortRetryIgnore; break}
   'YesNoCancel' {$btn = [System.Windows.Forms.MessageBoxButtons]::YesNoCancel; break}
   'YesNo' {$btn = [System.Windows.Forms.MessageBoxButtons]::yesno; break}
   'RetryCancel'{$btn = [System.Windows.Forms.MessageBoxButtons]::RetryCancel; break}
   default {$btn = [System.Windows.Forms.MessageBoxButtons]::RetryCancel; break}
}
 
  # Display the message box
  $Return=[System.Windows.Forms.MessageBox]::Show($Message,$Title,$btn)
  $Return
}


If ($Debug) { Write-Host "Importing the Config Manager Module --------- $(Get-Date)" -ForegroundColor Cyan }
# Import the ConfigurationManager.psd1 module 
if((Get-Module ConfigurationManager) -eq $null) {
    Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" @initParams 
}

If ($Debug) { Write-Host "Connecting to the Config Manager Site Drive ----------- $(Get-Date)" -ForegroundColor Cyan }
# Connect to the site's drive if it is not already present
if((Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue) -eq $null) {
    New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName @initParams
}

# Set the current location to be the site code.
Set-Location "$($SiteCode):\" @initParams

If ($Debug) { Write-Host "Running the query for all installed software at your location ----------- $(Get-Date)" -ForegroundColor Cyan }
# Pull all installed software from SCCM
$AllSoftware = Invoke-CMQuery -Name $QueryName

If ($CSVExists){
If ($Debug) { Write-Host "Importing the previous mappings ----------- $(Get-Date)" -ForegroundColor Cyan }
# Import Exisiting Mappings
$ExistingMappings = Import-Csv $ScriptPathParent\Applications.csv
}

If ($Debug) { Write-Host "Pulling all available applications from Config Manager and sorting them alphabetically ----------- $(Get-Date)" -ForegroundColor Cyan }
# Get all unique Application names and sort them alphabetically
$AvailableApps = Get-CMApplication -Fast | Select LocalizedDisplayName
$AvailableApps = $AvailableApps | Sort-Object -Property LocalizedDisplayName | Get-Unique -AsString

If ($Debug) { Write-Host "Matching up all the Installed Apps in the environment with the previous mapped ones ----------- $(Get-Date)" -ForegroundColor Cyan }
If ($CSVExists){
foreach ($Mapping in $ExistingMappings)
{
    $AllApplicationMappings += [pscustomobject]@{
            "DisplayName"="$($Mapping.DisplayName)"
            "NewApp" = $($Mapping.NewApp)
        }
}

}

foreach ($Software in $($AllSoftware | where-object {$_.ProductName -NotIn $AllApplicationMappings.DisplayName} ) ){
    $AllApplicationMappings += [pscustomobject]@{
        "DisplayName"="$($Software.ProductName)"
    }
}

$AllApplicationMappings = $AllApplicationMappings | Sort-Object -Property DisplayName

If ($Debug) { Write-Host "Creating the columns for the form --------- $(Get-Date)" -ForegroundColor Cyan }
# Datatable for your CSV content
$DataTable1 = New-Object System.Data.DataTable
[void] $DataTable1.Columns.Add("DisplayName")
[void] $DataTable1.Columns.Add("NewApp")

If ($Debug) { Write-Host "Creating the mappings for the columns ----------- $(Get-Date)" -ForegroundColor Cyan }
# Create Mappings according to current installed software list
$AllApplicationMappings | ForEach-Object {
    [void] $DataTable1.Rows.Add($($_.DisplayName), $($_.NewApp))
    }

If ($Debug) { Write-Host "Creating the mappings for the combo boxes ----------- $(Get-Date)" -ForegroundColor Cyan }
$DataTable2 = New-Object System.Data.DataTable
[void] $DataTable2.Columns.Add("NewApp")
[void] $DataTable2.Rows.Add("")
ForEach ($AvailableApp in $AvailableApps) {
    [void] $DataTable2.Rows.Add("$($AvailableApp.LocalizedDisplayName)")    
}
    

If ($Debug) { Write-Host "Configuring the form -----$(Get-Date) " -ForegroundColor Cyan }
# Form
$Form = New-Object System.Windows.Forms.Form
$Form.Size = New-Object System.Drawing.Size(710,700)
$Form.StartPosition = "CenterScreen"
$Form.Text = "Application Mapper Creation GUI"
$Form.Controls.Add($Label)
$Form.AutoSizeMode = "GrowAndShrink"
$Form.MinimizeBox = $False
$Form.MaximizeBox = $False
$Form.WindowState = "Normal"
$Form.ShowInTaskbar = $true
$Form.ShowIcon = $False

# Label
If ($Debug) { Write-Host "Configuring the Label -----$(Get-Date) " -ForegroundColor Cyan }
$Label = New-Object System.Windows.Forms.Label
$Label.Text = "Use this form to create the application mapping for dynamic installing applications. The left column lists all the inventoried software in Config Manager that has been installed on any managed device. The right column is the application you are going to replace or map it to."
$Label.Location = New-Object System.Drawing.Size(10,10)
$Label.Size = New-Object System.Drawing.Size(680,90)
$Label.Font = New-Object System.Drawing.Font("Microsoft Sans Serif",10,0,3,1)
$Form.Controls.Add($Label)

# Form event handlers
$Form.Add_Shown({
    $Form.Activate()
    })

# Datagridview
If ($Debug) { Write-Host "Configuring the DataGridView -----$(Get-Date) " -ForegroundColor Cyan }
$DGV = New-Object System.Windows.Forms.DataGridView
$DGV.Anchor = [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Top
$DGV.Location = New-Object System.Drawing.Size(10,100) 
$DGV.Size = New-Object System.Drawing.Size(680,500)
$DGV.Font = New-Object System.Drawing.Font("Microsoft Sans Serif",10,0,3,1)
$DGV.BackgroundColor = "#ffffffff"
$DGV.BorderStyle = "Fixed3D"
$DGV.AlternatingRowsDefaultCellStyle.BackColor = "#ffe6e6e6"
$DGV.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
$DGV.AutoSizeRowsMode = [System.Windows.Forms.DataGridViewAutoSizeRowsMode]::AllCells
$DGV.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$DGV.ClipboardCopyMode = "EnableWithoutHeaderText"
$DGV.AllowUserToOrderColumns = $True
$DGV.DataSource = $DataTable1
$DGV.AutoGenerateColumns = $False
$Form.Controls.Add($DGV)

# Datagridview columns
If ($Debug) { Write-Host "Configuring the DataGrdView Column 1 -----$(Get-Date) " -ForegroundColor Cyan }
$Column1 = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$Column1.Name = "DisplayName"
$Column1.HeaderText = "DisplayName"
$Column1.DataPropertyName = "DisplayName"
$Column1.AutoSizeMode = "Fill"

If ($Debug) { Write-Host "Configuring the DataGrdView Column 2 -----$(Get-Date) " -ForegroundColor Cyan }
$Column2 = New-Object System.Windows.Forms.DataGridViewComboBoxColumn
$Column2.Name = "NewApp"
$Column2.HeaderText = "NewApp"
$Column2.DataSource = $DataTable2
$Column2.ValueMember = "NewApp"
$Column2.DisplayMember = "NewApp"
$Column2.DataPropertyName = "NewApp"

$DGV.Columns.AddRange($Column1, $Column2)

# Button to export data
If ($Debug) { Write-Host "Configuring the Ok/Create CSV button -----$(Get-Date) " -ForegroundColor Cyan }
$Button = New-Object System.Windows.Forms.Button
$Button.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
$Button.Location = New-Object System.Drawing.Size(10,620) 
$Button.Text = "Create CSV"
$Button.Size = New-Object System.Drawing.Size(75,23)
$Button.DialogResult = [System.Windows.Forms.DialogResult]::Ok
$Form.AcceptButton = $Button
$Form.Controls.Add($Button)

# Cancel Button
If ($Debug) { Write-Host "Configuring the Cancel Button -----$(Get-Date) " -ForegroundColor Cyan }
$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Point(610,620)
$CancelButton.Size = New-Object System.Drawing.Size(75,23)
$CancelButton.Text = 'Cancel'
$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $CancelButton
$form.Controls.Add($CancelButton)

<#
# Button event handlers
$Button.Add_Click({
    $DataToExport = $DataTable1 | Where { -not [string]::IsNullOrEmpty($_.NewApp) }
    $DataToExport | Export-Csv $ScriptPathParent\Applications.csv -NoTypeInformation -Force
    msgbox -Message "You have successfully created the updated csv file for mapping applications." -Title "Success in Creating CSV" -buttons ok
    })
#>

If ($Debug) { Write-Host "Showing the form ------------ $(Get-Date)" -ForegroundColor Cyan }
# Show form
$result = $Form.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK)
{
    $DataToExport = $DataTable1 | Where { -not [string]::IsNullOrEmpty($_.NewApp) }
    $DataToExport | Export-Csv $ScriptPathParent\Applications.csv -NoTypeInformation -Force
    msgbox -Message "You have successfully created the updated csv file for mapping applications." -Title "Success in Creating CSV" -buttons ok
}