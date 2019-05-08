<#
    .SYNOPSIS
    Queries currently installed applications and maps currently installed apps to fresh, updated apps for deployment
    .DESCRIPTION
    This script will query the current computer configuration for all installed apps, import the csv file containing the mappings between previously installed apps
    and the new apps that will replace the currently installed ones for a smoother, more dynamic, refresh scenario.
    .EXAMPLE
    .\Set-DynamicApps.ps1
    .NOTES
    FileName:    Set-DynamicApps.ps1
    Author:      John Yoakum
    Created:     2019-03-01
    
    Version history:
    1.0.0 - (2019-03-01) Script created
    2.0.0 - (2019-05-07) Added more logging, set for debugging options, updated and verified variables

#>
# Stores the full path to the parent directory of this powershell script
# e.g. C:\Scripts\GoogleApps
$ScriptPathParent = split-path -Parent -Path $MyInvocation.MyCommand.Definition

$Debug = $true

# Create Archive folder if it doesn't exist for this log file and for sending logs
$TestPath = Test-Path -path c:\Archive
If (!$TestPath) { $LogsDirectory = New-Item -Path 'c:\Archive' -Force -ItemType Directory }
Else { $LogsDirectory = 'C:\Archive' }

# Initialize the TS Environment
If (!$Debug) { $tsenv = New-Object -COMObject Microsoft.SMS.TSEnvironment }

$InstalledApps = @()

# Function to write to log file
function Write-CMLogEntry {
    param (
        [parameter(Mandatory = $true, HelpMessage = 'Value added to the log file.')]
        [ValidateNotNullOrEmpty()]
        [string]$Value,
        [parameter(Mandatory = $true, HelpMessage = 'Severity for the log entry. 1 for Informational, 2 for Warning and 3 for Error.')]
        [ValidateNotNullOrEmpty()]
        [ValidateSet('1', '2', '3')]
        [string]$Severity,
        [parameter(Mandatory = $false, HelpMessage = 'Name of the log file that the entry will written to.')]
        [ValidateNotNullOrEmpty()]
        [string]$FileName = 'PackageMapping.log'
    )
    # Determine log file location
    $LogFilePath = Join-Path -Path $LogsDirectory -ChildPath $FileName
		
    # Construct time stamp for log entry
    $Time = -join @((Get-Date -Format 'HH:mm:ss.fff'), '+', (Get-WmiObject -Class Win32_TimeZone | Select-Object -ExpandProperty Bias))
		
    # Construct date for log entry
    $Date = (Get-Date -Format 'MM-dd-yyyy')
		
    # Construct context for log entry
    $Context = $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)
		
    # Construct final log entry
    $LogText = "<![LOG[$($Value)]LOG]!><time=""$($Time)"" date=""$($Date)"" component=""PackageMapping"" context=""$($Context)"" type=""$($Severity)"" thread=""$($PID)"" file="""">"
		
    # Add value to log file
    try {
        Out-File -InputObject $LogText -Append -NoClobber -Encoding Default -FilePath $LogFilePath -ErrorAction Stop
    }
    catch {
        Write-Warning -Message "Unable to append log entry to PackageMapping.log file. Error message at line $($_.InvocationInfo.ScriptLineNumber): $($_.Exception.Message)"
    }
}

# Function for Special Application Mapping
Function Get-SpecialApplication() {
    [CmdletBinding()]
    param($ApplicationToPass)
    $SpecialApplication = $ApplicationMapping | Where-Object { ($_.DisplayName -like $ApplicationToPass) -and ($AllInstalledApps.DisplayName -like $ApplicationToPass ) } | Get-Unique -AsString
    If ( $SpecialApplication.DisplayName -like $ApplicationToPass ) {
        Return $SpecialApplication
    }
}

# Get list of apps

# Add list of 32-bit apps on a 64-bit system
if ( [Environment]::Is64BitOperatingSystem -eq $true ) {
    Write-CMLogEntry -Value 'Getting list of previously installed 32-bit apps, on 64-bit system (WoW)' -Severity 1
    $InstalledApps += @(Get-ItemProperty -Path HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName)
}

# Add list of 64-bit apps on a 64-bit system, or 32-bit apps on a 32-bit system
Write-CMLogEntry -Value "Getting list of previously installed apps (non-WoW)" -Severity 1
$InstalledApps += @(Get-ItemProperty -Path HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName)

# Sort the list of apps and only pull unique values and log them all
$AllInstalledApps = $InstalledApps | Sort-Object -Property DisplayName | Get-Unique -AsString
Write-CMLogEntry -Value 'Listed below are the previously installed apps' -Severity 1
ForEach ($InstalledApplication in $AllInstalledApps) {
    Write-CMLogEntry -Value "Previous Application: $($InstalledApplication.DisplayName)" -Severity 1
}

Write-CMLogEntry -Value "Exporting CSV for All Previously Installed Apps. Stored at $LogsDirectory\PreviouslyInstalledApps.csv" -Severity 1

# Save a copy of all the installed apps on the local hard drive
$AllInstalledApps | Export-CSV -Path $LogsDirectory\PreviouslyInstalledApps.csv -NoTypeInformation -Encoding UTF8

Write-CMLogEntry -Value 'Importing CSV for application mapping.' -Severity 1

# Import the csv file with the mappings
$ApplicationMapping = Import-Csv -Path $ScriptPathParent\Applications.csv
Write-CMLogEntry -Value 'Matching Apps for replacement.' -Severity 1

# Create variable and populate variable for the matching applications between the CSV and Installed Apps
$MatchingApps = @()

# Pulls all firefoxes if installed
$HasFirefox = $ApplicationMapping | Where-object { (($_.DisplayName -like 'Mozilla Firefox *(*') -and ($AllInstalledApps.DisplayName -like 'Mozilla Firefox *(*')) -and (($_.DisplayName -notlike 'Mozilla Firefox * ESR *') -and ($AllInstalledApps.DisplayName -notlike 'Mozilla Firefox * ESR *')) } | Get-Unique -AsString
If ( ($HasFirefox.DisplayName -like 'Mozilla Firefox *') -and ($HasFirefox -notlike 'Mozilla Firefox * ESR *') ) {
    $MatchingApps += $HasFirefox
}

$MatchingApps += $ApplicationMapping | Where-Object { $AllInstalledApps.DisplayName -contains $_.DisplayName } | Get-Unique -AsString

# Special Mappings for versioned apps
$SpecialApplicationMapping = 'Mozilla Firefox * ESR *', 'Microsoft Office *', 'Audacity *', 'Google Chrome*', '7-Zip *', 'Evernote *', 'FileZilla*', 'FileOpen Client *', 'Adobe Shockwave Player *', `
    'Adobe Reader*', 'ArcGIS * for Desktop*', 'Camstudio*', 'Cisco Packet Tracer *', 'Firefox Developer *', 'Gephi *', 'GIMP*', 'HandBrake *', 'Inkscape *', 'KeePass*', 'Labstats*', `
    'LibreOffice *', 'Logger Pro *', 'Mendeley Desktop *', 'Microsoft SQL Server Management *', 'Mozilla Thunderbird *', 'NetLogo *', 'Octave*', 'OpenOffice*', `
    'Opera *', 'Oracle VM Virtual *', 'Paint.NET*', 'PuTTY *', 'Python *', 'R for Windows *', 'SQL Server * Management *', 'Toad Data Point *', 'VLC media player*', 'VUE*', 'WinSCP*', 'Wireshark*', `
    'Wolfram CDF *', 'Wolfram Mathematica *', 'WolfVision *', 'Z+FLaserControl *', 'XShell *', 'Jave* Update *', 'Adobe Digital *', 'Adobe Flash Player * N*', 'Adobe Flash Player * P*', 'Adobe Flash Player* A*', `
    'Camtasia*', 'CloudCompare*', 'Corpscon*', 'Counterpointer*', 'gnuplot*', 'Google Earth*', 'ImageMagick *', 'Jave* Update *(*', 'Skype*', 'Graphical Analysis*', 'NetBeans IDE*', 'EndNote*', 'Git *', 'Focusky *', 'JetBrains PyCharm *', `
    'Nmap*', 'TortoiseGit*', 'TortoiseHg*', 'TortoiseSVN*', 'Symantec*'
ForEach ($SpecialApplication in $SpecialApplicationMapping) {
    $MatchingApps += Get-SpecialApplication -ApplicationToPass $SpecialApplication
}

# Sort the Apps and get unique values
$MatchingApps = $MatchingApps | Sort-Object -Property DisplayName | Get-Unique -AsString

# Assign the Unique Apps to be installed to a specific variable
$ApplicationsToInstall = $MatchingApps.NewApp | Get-Unique -AsString

$Apps = $ApplicationsToInstall |Sort-Object | Get-Unique -AsString

#Reset initial count to 0
$Count = 0

# Section of code to set the task sequence value for the newly installed apps
foreach ($ApplicationName in $Apps) {
    $Id = "{0:D2}" -f $Count
    $AppId = "UAAapps$Id" 
    If (!$Debug) {$TSEnv.Value($AppId) = $ApplicationName}
    Write-CMLogEntry -Value "Task Sequence Variable Name: $AppID     -----   Application: $ApplicationName " -Severity 1
    $Count = $Count + 1
    
}

# Add additional Apps - Teams
$Id = "{0:D2}" -f $Count
$AppId = "UAAapps$Id" #You can make the base variable anything you would like as long as you reference it in the Install Application task sequence step
If (!$Debug) { $TSEnv.Value($AppId) = "Teams Machine-Wide Installer" }
Write-CMLogEntry -Value "Task Sequence Variable Name: $AppID     -----   Application: Teams Machine-Wide Installer " -Severity 1
# Uncomment out the following line if you continue to adding additional apps for dynamic installation 
#$Count = $Count + 1

<# Use the following section to continue to add additional apps that were not included through the process above
# Add additional Apps - OneDrive
$Id = "{0:D2}" -f $Count
$AppId = "UAAapps$Id"
If (!$Debug) { $TSEnv.Value($AppId) = "Microsoft OneDrive for Business Client" }
Write-CMLogEntry -Value "Task Sequence Variable Name: $AppID     -----   Application: Teams Machine-Wide Installer " -Severity 1
$Count = $Count + 1
#>