<#
.SYNOPSIS
    Generates an audit report for Configuration Manager.
.DESCRIPTION
    The Get-CMAuditReport.ps1 script retrieves various audit data from a Configuration Manager environment
    and generates a report in Word format. The report includes information about collections, applications,
    packages, deployments, task sequences, site details, site servers, and SQL server details.
.PARAMETER None
    This script does not accept any parameters.
.OUTPUTS
    Word Document
    The script generates a Word document containing the audit report.
.EXAMPLE
    .\Get-CMAuditReport.ps1
    Generates an audit report for the Configuration Manager environment and saves it as a Word document.
.NOTES
    - This script requires the Configuration Manager PowerShell module to be installed.
    - The script should be run on a system with the Configuration Manager console installed.
    - The Word.Application COM object is used to automate Word, so Word needs to be installed on the system.
.LINK
    [Configuration Manager Documentation](https://docs.microsoft.com/en-us/mem/configmgr/)
.NOTES
    Version: 2.0
    Creation Date: 2023-06-10
    Copyright (c) 2023 https://github.com/bentman
    https://github.com/bentman/PoShConfigManAuditReport
#>

#Requires -Module ($env:SMS_ADMIN_UI_PATH.Substring(0,$env:SMS_ADMIN_UI_PATH.Length-5) + '\ConfigurationManager.psd1')
cd ((Get-PSDrive -PSProvider CMSite).Name + ':')

# Gather the necessary audit data
$Site = Get-CMSite
$SiteServers = Get-CMSiteSystem
$SqlServer = $Site.DatabaseServerName
$SqlDb = $Site.DatabaseName
$Collections = Get-CMDeviceCollection
$Applications = Get-CMApplication
$Packages = Get-CMPackage
$Deployments = Get-CMDeployment
$TaskSequences = Get-CMTaskSequence | ForEach-Object {
    $TSSteps = Get-CMTaskSequenceStep -TaskSequencePackageID $_.PackageID
    [PSCustomObject]@{
        "TaskSequence" = $_
        "Steps" = $TSSteps
    }
}

# Start Word and create a new document
$word = New-Object -ComObject Word.Application
$word.Visible = $true
$doc = $word.Documents.Add()
$selection = $word.Selection

# Write the audit data to the Word document with better formatting
$selection.Font.Bold = 1
$selection.Font.Size = 16
$selection.TypeText("Config Manager Audit Data")
$selection.TypeParagraph()

# Add Site
$selection.Font.Bold = 1
$selection.Font.Size = 14
$selection.TypeText("Site")
$selection.TypeParagraph()
$selection.Font.Bold = 0
$selection.Font.Size = 12
$selection.TypeText("Name: " + $Site.Name)
$selection.TypeParagraph()
$selection.TypeText("Site Code: " + $Site.SiteCode)
$selection.TypeParagraph()

# Add Site Servers
$selection.Font.Bold = 1
$selection.Font.Size = 14
$selection.TypeText("Site Servers")
$selection.TypeParagraph()
foreach ($server in $SiteServers) {
    $selection.Font.Bold = 0
    $selection.Font.Size = 12
$selection.TypeText("Server Name: " + $server.ServerName)
$selection.TypeParagraph()
$selection.TypeText("Roles: " + ($server.Roles | Out-String))
$selection.TypeParagraph()
}

# Add SQL Server Details
$selection.Font.Bold = 1
$selection.Font.Size = 14
$selection.TypeText("SQL Server Details")
$selection.TypeParagraph()
$selection.Font.Bold = 0
$selection.Font.Size = 12
$selection.TypeText("SQL Server: " + $SqlServer)
$selection.TypeParagraph()
$selection.TypeText("SQL Database: " + $SqlDb)
$selection.TypeParagraph()

# Add Collections
$selection.Font.Bold = 1
$selection.Font.Size = 14
$selection.TypeText("Collections")
$selection.TypeParagraph()
foreach ($collection in $Collections) {
    $selection.Font.Bold = 0
    $selection.Font.Size = 12
    $selection.TypeText("Name: " + $collection.Name)
    $selection.TypeParagraph()
    $selection.TypeText("CollectionID: " + $collection.CollectionID)
    $selection.TypeParagraph()
    $selection.TypeText("MemberCount: " + $collection.MemberCount)
    $selection.TypeParagraph()
}

# Add Applications
$selection.Font.Bold = 1
$selection.TypeText("Applications")
$selection.TypeParagraph()
foreach ($application in $Applications) {
    $selection.Font.Bold = 0
    $selection.TypeText("Name: " + $application.LocalizedDisplayName)
    $selection.TypeParagraph()
    $selection.TypeText("ApplicationID: " + $application.CI_ID)
    $selection.TypeParagraph()
    $selection.TypeText("Version: " + $application.SoftwareVersion)
    $selection.TypeParagraph()
}

# Add Packages
$selection.Font.Bold = 1
$selection.TypeText("Packages")
$selection.TypeParagraph()
foreach ($package in $Packages) {
    $selection.Font.Bold = 0
    $selection.TypeText("Name: " + $package.Name)
    $selection.TypeParagraph()
    $selection.TypeText("PackageID: " + $package.PackageID)
    $selection.TypeParagraph()
    $selection.TypeText("Version: " + $package.PackageVersion)
    $selection.TypeParagraph()
}

# Add Deployments
$selection.Font.Bold = 1
$selection.TypeText("Deployments")
$selection.TypeParagraph()
foreach ($deployment in $Deployments) {
    $selection.Font.Bold = 0
    $selection.TypeText("PackageID: " + $deployment.PackageID)
    $selection.TypeParagraph()
    $selection.TypeText("Target: " + $deployment.TargetCollectionID)
    $selection.TypeParagraph()
    $selection.TypeText("State: " + $deployment.State)
    $selection.TypeParagraph()
    $selection.TypeText("StartTime: " + $deployment.StartTime)
    $selection.TypeParagraph()
    $selection.TypeText("EndTime: " +```powershell
$deployment.EnforcementDeadline)
$selection.TypeParagraph()
}

# Add Task Sequences
$selection.Font.Bold = 1
$selection.TypeText("Task Sequences")
$selection.TypeParagraph()
foreach ($ts in $TaskSequences) {
    $selection.Font.Bold = 0
    $selection.TypeText("Name: " + $ts.TaskSequence.Name)
    $selection.TypeParagraph()
    $selection.TypeText("PackageID: " + $ts.TaskSequence.PackageID)
    $selection.TypeParagraph()
    $selection.Font.Bold = 1
    $selection.TypeText("Steps")
    $selection.TypeParagraph()
    foreach ($step in $ts.Steps) {
        $selection.Font.Bold = 0
        $selection.TypeText("Step: " + $step.Name)
        $selection.TypeParagraph()
        $selection.TypeText("Type: " + $step.StepType)
        $selection.TypeParagraph()
    }
    $selection.TypeParagraph()
}

# Save and close the document
$doc.SaveAs([Ref] "C:\Temp\AuditData.docx", [Ref] 16) # 16 = wdFormatDocumentDefault
$doc.Close()
