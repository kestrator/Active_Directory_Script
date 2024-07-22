<#
.SYNOPSIS
    The Create-PingCastleReport.ps1 script launches the PingCastle analysis with the default parameters and sends the report by mail.

.DESCRIPTION
    The Create-PingCastleReport.ps1 script allows you to obtain information on the Active Directory security score with the PingCastle executable.
    You will receive the report by mail.
    You can configure this script in a scheduled task to obtain a periodic report.


.EXAMPLE
    Create-PingCastleReport.ps1

    This script generates a report and sends it by e-mail.


.NOTES
    NAME:    Create-PingCastleReport.ps1
    AUTHOR:  Jonathan BAUZONE JBE
    DATE:    2024/07/22
    WWW:     N/A
    TWITTER: @BznJnthn

    VERSION HISTORY:
    1.0 2024.07.22
        Initial Version
#>

$ErrorActionPreference = 'Stop'
$InformationPreference = 'Continue'
#Specify domain name
$Domains = ""
#Specify script location
$PSScriptRoot = ""

#region Variable
foreach ($domain in $Domains)
{
    $ApplicationName = 'PingCastle'
    $PingCastle = [pscustomobject]@{
        Name            = $ApplicationName
        ProgramPath     = Join-Path $PSScriptRoot $ApplicationName
        ProgramName     = '{0}.exe' -f $ApplicationName
        Arguments       = "--healthcheck --server $($Domain) --level Full"
        ReportFileName  = 'ad_hc_{0}' -f $domain.ToLower()
        ReportFolder    = "Reports"
        OldFolder       = "Old"
        ScoreFileName   = '{0}Score.txt' -f $ApplicationName
        ProgramUpdate   = '{0}AutoUpdater.exe' -f $ApplicationName
        ArgumentsUpdate = '--wait-for-days 30'
        
    }
             

    $pingCastleFullpath = Join-Path $PingCastle.ProgramPath $PingCastle.ProgramName
    $pingCastleUpdateFullpath = Join-Path $PingCastle.ProgramPath $PingCastle.ProgramUpdate
    $pingCastleReportLogs = Join-Path $PingCastle.ProgramPath $PingCastle.ReportFolder
    $pingCastleReportsOld = Join-Path $pingCastleReportLogs $PingCastle.OldFolder
    $PingCastleScoreFileName = $domain + $PingCastle.ScoreFileName
    $pingCastleScoreFileFullpath = Join-Path $pingCastleReportLogs $PingCastleScoreFileName
    $pingCastleReportFullpath = Join-Path $PingCastle.ProgramPath ('{0}.html' -f $PingCastle.ReportFileName)
    $pingCastleReportXMLFullpath = Join-Path $PingCastle.ProgramPath ('{0}.xml' -f $PingCastle.ReportFileName)
    $pingCastleReportDate = Get-Date -UFormat %Y%m%d_%H%M%S
    $pingCastleReportFileNameDate = ('{0}_{1}' -f $pingCastleReportDate, ('{0}.html' -f $PingCastle.ReportFileName))

#$sentNotification = $true

    $splatProcess = @{
        WindowStyle = 'Hidden'
        Wait        = $true
    }
    #endregion

    # Check if program exist
    if (-not(Test-Path $pingCastleFullpath)) {
        Write-Error -Message ("Path not found {0}" -f $pingCastleFullpath)
    }

# Check if log directory exist. If not, create it
    if (-not (Test-Path $pingCastleReportLogs)) {
        try {
            $null = New-Item -Path $pingCastleReportLogs -ItemType directory
        }
        Catch {
            Write-Error -Message ("Error for create directory {0}" -f $pingCastleReportLogs)
        }
    }

# Check if Old directory exist. If not, create it
    if (-not (Test-Path $pingCastleReportsOld)) {
        try {
            $null = New-Item -Path $pingCastleReportsOld -ItemType directory
        }
        Catch {
            Write-Error -Message ("Error for create directory {0}" -f $pingCastleReportsOld)
        }
    }

# Try to start program and catch any error
    try {
        Set-Location -Path $PingCastle.ProgramPath
        Start-Process -FilePath $pingCastleFullpath -ArgumentList $PingCastle.Arguments @splatProcess
    }

    Catch {
        Write-Error -Message ("Error for execute {0}" -f $pingCastleFullpath)
    }

# Check if report exist after execution
    foreach ($pingCastleTestFile in ($pingCastleReportFullpath, $pingCastleReportXMLFullpath)) {
        if (-not (Test-Path $pingCastleTestFile)) {
            Write-Error -Message ("Report file not found {0}" -f $pingCastleTestFile)
        }
    }

# Get content on XML file
    try {
        $contentPingCastleReportXML = $null
        $contentPingCastleReportXML = (Select-Xml -Path $pingCastleReportXMLFullpath -XPath "/HealthcheckData").node
    }
    catch {
        Write-Error -Message ("Unable to read the content of the xml file {0}" -f $pingCastleReportXMLFullpath)
    }

# Convert to json all score from xml file
    try {
        $contentPingCastleReport = $contentPingCastleReportXMLToJSON = $null
        $contentPingCastleReport = $contentPingCastleReportXML | Select-Object *Score
        $contentPingCastleReportXMLToJSON = $contentPingCastleReport | ConvertTo-Json -Compress
    }
    catch {
        Write-Error -Message "Unable to convert the content to json"
    }
# Update report file with current score
    try {
        $contentPingCastleReportXMLToJSON | Out-File $pingCastleScoreFileFullpath -Force
    }
    Catch {
        Write-Error -Message ("Error for update report score file {0}" -f $pingCastleScoreFileFullpath)
    }


# Move report to logs directory
    try {
        $pingCastleMoveFile = (Join-Path $pingCastleReportLogs $pingCastleReportFileNameDate)
        Move-Item -Path $pingCastleReportFullpath -Destination $pingCastleMoveFile
    }
    catch {
        Write-Error -Message ("Error for move report file to logs directory {0}" -f $pingCastleReportFullpath)
    }

# Try to start update program and catch any error
    try {
        Write-Information "Trying to update"
        Start-Process -FilePath $pingCastleUpdateFullpath -ArgumentList $PingCastle.ArgumentsUpdate @splatProcess
        Write-Information "Update completed"
    }
    Catch {
        Write-Error -Message ("Error for execute update program {0}" -f $pingCastleUpdateFullpath)
    }

#Get Date
    $ReportDate = Get-Date -format "dd-MM-yyyy"

#Configuration Variables for E-mail
    $SmtpServer = "" #or IP Address such as "10.125.150.250"
    $EmailFrom = "ReportPingcastle no-reply@"
    $EmailTo = ""
    $EmailSubject = "PingCastle "+$Domain+ " - Rapport Hebdomadaire "+$ReportDate

#HTML Template
    $EmailBody = @"
    <table style="width: 570px" style="border-collapse: collapse; border: 1px solid #008080;">
     <tr>
        <td style="height: 39px;"><img src="data:image/png;base64,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"></td>
        <td colspan="2" style="font-size: large; height: 39px">PingCastle DomainName - Rapport Hebdomadaire VarReportDate </td>
     </tr>
     <tr>
        <td colspan="3" bgcolor="orange" style="color: #FFFFFF; font-size: large; height: 2px; text-align: center;"></td>
     </tr>
     <tr style="border-bottom-style: solid; border-bottom-width: 1px; padding-bottom: 1px">
        <td style="width: 200px; height: 39px">Score Global</td>
        <td style="text-align: center; height: 39px; width: 75px"><b>GLobalScore</b></td>
        <td style="height: 39px; color:ColorResult1"><progress style="color:ColorResult1" value="GBScore" max="100">GBScore points</progress></td>
     </tr>
      <tr style="height: 39px; border: 1px solid #008080">
        <td style="width: 200px; height: 39px; padding:0 0 0 20px"> Objets obsolètes</td>
        <td style="text-align: center; height: 39px; width: 75px"><b>StaleObjects</b></td>
        <td style="height: 39px; color:ColorResult2"><progress style="color:ColorResult2" value="StaleScore" max="100">StaleScore points</progress></td>
     </tr>
     <tr style="height: 39px; border: 1px solid #008080">
        <td style="width: 200px; height: 39px; padding:0 0 0 20px"> Comptes à privilèges</td>
        <td style="text-align: center; height: 39px; width: 75px"><b>AccountPrivileged</b></td>
        <td style="height: 39px; color:ColorResult3"><progress style="color:ColorResult3" value="AccPriScore" max="100">AccPriScore points</progress></td>
     </tr>
     <tr style="height: 39px; border: 1px solid #008080">
        <td style="width: 200px; height: 39px; padding:0 0 0 20px"> Relations d'approbation</td>
        <td style="text-align: center; height: 39px; width: 75px"><b>Trusts</b></td>
        <td style="height: 39px; color:ColorResult4"><progress style="color:ColorResult4" value="TrsScore" max="100">TrsScore points</progress></td>
     </tr>
     <tr style="height: 39px; border: 1px solid #008080">
        <td style="width: 200px; height: 39px; padding:0 0 0 20px"> Anomalies</td>
        <td style="text-align: center; height: 39px; width: 75px"><b>Anomaly</b></td>
        <td style="height: 39px; color:ColorResult5"><progress style="color:ColorResult5" value="AnoScore" max="100">AnoScore points</progress></td>
     </tr>
    </table>
"@

#Get Values for Approved & Rejected variables
    $GlobalScore = $contentPingCastleReport.GlobalScore
    $StaleObjectsScore= $contentPingCastleReport.StaleObjectsScore
    $PrivilegiedGroupScore=$contentPingCastleReport.PrivilegiedGroupScore
    $TrustScore=$contentPingCastleReport.TrustScore
    $AnomalyScore=$contentPingCastleReport.AnomalyScore

    if ($GlobalScore -le 25){
        $GlobalColor = "green"
    }

    elseif ($GlobalScore -gt 25 -and $GlobalScore -le 50){
        $GlobalColor = "yellow"
    }
    elseif ($GlobalScore -gt 50 -and $GlobalScore -le 75){
        $GlobalColor = "orange"
    }
    else{
        $GlobalColor = "red"
    }

    if ($StaleObjectsScore -le 25){
        $StaleColor = "green"
    }
    elseif ($StaleObjectsScore -gt 25 -and $StaleObjectsScore -le 50){
        $StaleColor = "yellow"
    }
    elseif ($StaleObjectsScore -gt 50 -and $StaleObjectsScore -le 75){
        $StaleColor = "orange"
    }
    else{
        $StaleColor = "red"
    }

    if ($PrivilegiedGroupScore -le 25){
        $PrivColor = "green"
    }
    elseif ($PrivilegiedGroupScore -gt 25 -and $PrivilegiedGroupScore -le 50){
        $PrivColor = "yellow"
    }
    elseif ($PrivilegiedGroupScore -gt 50 -and $PrivilegiedGroupScore -le 75){
        $PrivColor = "orange"
    }
    else{
        $PrivColor = "red"
    }

    if ($TrustScore -le 25){
        $TrustColor = "green"
    }
    elseif ($TrustScore -gt 25 -and $TrustScore -le 50){
        $TrustColor = "yellow"
    }
    elseif ($TrustScore -gt 50 -and $TrustScore -le 75){
        $TrustColor = "orange"
    }
    else{
        $TrustColor = "red"
    }

    if ($AnomalyScore -le 25){
        $AnoColor = "green"
    }
    elseif ($AnomalyScore -gt 25 -and $AnomalyScore -le 50){
        $AnoColor = "yellow"
    }
    elseif ($AnomalyScore -gt 50 -and $AnomalyScore -le 75){
        $AnoColor = "orange"
    }
    else{
        $AnoColor = "red"
    }

    #Replace the Variables VarApproved, VarRejected and VarReportDate
    $EmailBody= $EmailBody.Replace("VarReportDate",$ReportDate)
    $EmailBody= $EmailBody.Replace("DomainName",$Domain)
    $EmailBody= $EmailBody.Replace("GLobalScore",$GlobalScore)
    $EmailBody= $EmailBody.Replace("GBScore",$GlobalScore)
    $EmailBody= $EmailBody.Replace("ColorResult1",$GlobalColor)
    $EmailBody= $EmailBody.Replace("StaleObjects",$StaleObjectsScore)
    $EmailBody= $EmailBody.Replace("StaleScore",$StaleObjectsScore)
    $EmailBody= $EmailBody.Replace("ColorResult2",$StaleColor)
    $EmailBody= $EmailBody.Replace("AccountPrivileged",$PrivilegiedGroupScore)
    $EmailBody= $EmailBody.Replace("AccPriScore",$PrivilegiedGroupScore)
    $EmailBody= $EmailBody.Replace("ColorResult3",$PrivColor)
    $EmailBody= $EmailBody.Replace("Trusts",$TrustScore)
    $EmailBody= $EmailBody.Replace("TrsScore",$TrustScore)
    $EmailBody= $EmailBody.Replace("ColorResult4",$TrustColor)
    $EmailBody= $EmailBody.Replace("Anomaly",$AnomalyScore)
    $EmailBody= $EmailBody.Replace("AnoScore",$AnomalyScore)
    $EmailBody= $EmailBody.Replace("ColorResult5",$AnoColor)

    #Send E-mail from PowerShell script
    Send-MailMessage -To $EmailTo -From $EmailFrom -Subject $EmailSubject -Body $EmailBody -BodyAsHtml -SmtpServer $SmtpServer -Attachments $pingCastleMoveFile -Encoding UTF8

    #Add-Content $report_html $EmailBody
    try {
        $pingCastleMove2File = (Join-Path $pingCastleReportsOld $pingCastleReportFileNameDate)
        $NewScoreFileName = $ReportDate + $PingCastleScoreFileName
        $pingCastleMove3File = (Join-Path $pingCastleReportsOld $NewScoreFileName)
        Move-Item -Path $pingCastleMoveFile -Destination $pingCastleMove2File
        Move-Item -Path $pingCastleScoreFileFullpath -Destination $pingCastleMove3File
    }
    catch {
        Write-Error -Message ("Error for move report file to Old directory {0}" -f $pingCastleReportFullpath)
    }
}
