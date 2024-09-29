[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory = $true)]
    [string]$InputServer
)
#endregion Paramsection

#Script starts here
$log = @()
StartScript -Log ([ref]$log)


try {
    $Services = Get-Service -ComputerName $InputServer | Where-Object { $_.DisplayName -Like "*Exchange*" } | Select-Object Status, Name, DisplayName
}
catch {
    $srxEnv.ResultMessage += "Failed to get services from $($InputServer)`n"
    exit 1
}

Write-Host "---------------------------------------------------------------"
Write-Host "Services gathered from server......!"

#Email Send
$EmailFrom = "PlatformService_Exchange_TestAutomation@brose.com"
#$EmailTo = @("rushiraj.gawande@brose.com", "anup.joshi@brose.com")
$EmailTo = @("rushiraj.gawande@brose.com", "anup.joshi@brose.com", "ralph.friedrich@brose.com", "max.herold@brose.com","pranay.bagul.temp@brose.com")
#Date
$currentTime = Get-Date -format "dd-MMM-yyyy HH:mm:ss"
$Subject = "Exchange Services Health and Port Status Report - $InputServer"
#Body
$style = @'
<style>body{font-family:`"Calibri`",`"sans-serif`"; font-size: 14px;}
@font-face
       {font-family:Tahoma;
       panose-1:2 11 6 4 3 5 4 4 2 4;}
       table{border: 1px solid black; border-collapse:collapse; mso-table-lspace:0pt; mso-table-rspace:0pt;}
       th{border: 1px solid black; background: #dddddd; padding: 5px; }
       td{border: 1px solid black; padding: 5px; }
       </style>
'@

#Services HTML
$mailbody += '<html><head><meta http-equiv=Content-Type content="text/html; charset=utf-8">' + $style + '</head><body>'
$mailbody += '<p>'
$mailbody += '-' * 100 + "<br />"
$mailbody += "Script run Date:" + '<b>' + " $currentTime" + '</b>' + "<br />"
$mailbody += '-' * 100 + "<br />"
$mailbody += "Services Health Check for the Server " + '<u>' + '<b>' + "$InputServer" + '</u>' + '</b>' + " is as below:" + "<br />"
$mailbody += '-' * 100
$mailbody += '</p>'

$mailbody += '<p><table>'
$mailbody += '<th>DisplayName</th><th>Name</th><th>Current Status</th><th>Status Required</th>'

$MatchStatusToServices = @{
    "HostControllerService"         = "Running";
    "MSComplianceAudit"             = "Running";
    "MSExchangeADTopology"          = "Running";
    "MSExchangeAntispamUpdate"      = "Running";
    "MSExchangeCompliance"          = "Running";
    "MSExchangeDagMgmt"             = "Running";
    "MSExchangeDelivery"            = "Running";
    "MSExchangeDiagnostics"         = "Running";
    "MSExchangeEdgeSync"            = "Running";
    "MSExchangeFastSearch"          = "Running";
    "MSExchangeFrontEndTransport"   = "Running";
    "MSExchangeHM"                  = "Running";
    "MSExchangeHMRecovery"          = "Running";
    "MSExchangeImap4"               = "Stopped";
    "MSExchangeIMAP4BE"             = "Stopped";
    "MSExchangeIS"                  = "Running";
    "MSExchangeMailboxAssistants"   = "Running";
    "MSExchangeMailboxReplication"  = "Running";
    "MSExchangeMitigation"          = "Running";
    "MSExchangeNotificationsBroker" = "Stopped";
    "MSExchangePop3"                = "Stopped";
    "MSExchangePOP3BE"              = "Stopped";
    "MSExchangeRepl"                = "Running";
    "MSExchangeRPC"                 = "Running";
    "MSExchangeServiceHost"         = "Running";
    "MSExchangeSubmission"          = "Running";
    "MSExchangeThrottling"          = "Running";
    "MSExchangeTransport"           = "Running";
    "MSExchangeTransportLogSearch"  = "Running";
    "MSExchangeUM"                  = "Running";
    "MSExchangeUMCR"                = "Running";
    "SearchExchangeTracing"         = "Stopped";
    "vmickvpexchange"               = "Stopped";
    "wsbexchange"                   = "Stopped";
}
#Comparing services so that we can ceheck and compare real status services also get to know if new services added
try {
    $Comparison = Compare-Object -ReferenceObject $MatchStatusToServices.Values -DifferenceObject $Services.Name
}
catch {
    $srxEnv.ResultMessage += "Error in comparing current and desired service status.`n"
    exit 2
}

if (![string]::IsNullOrEmpty($Comparison)) {
    foreach ($object in $Services) {
        <#if ($object.Status -eq "Running") {
            $mailbody += "<tr><td>" + $object.DisplayName + "</td><td>" + $object.Name + "</td><td>" + '<font color="green">' + $object.Status + '</font>' + "</td><td>" + if ($($MatchStatusToServices[$object.Name]) -eq "Running") { '<font color="green">' + $($MatchStatusToServices[$object.Name]) + "</font>" + "</td></tr>" }else { '<font color="red">' + $($MatchStatusToServices[$object.Name]) + "</font>" + "</td></tr>" }
        }
        else {
            $mailbody += "<tr><td>" + $object.DisplayName + "</td><td>" + $object.Name + "</td><td>" + '<font color="red">' + $object.Status + '</font>' + "</td><td>" + if ($($MatchStatusToServices[$object.Name]) -eq "Running") { '<font color="green">' + $($MatchStatusToServices[$object.Name]) + "</font>" + "</td></tr>" }else { '<font color="red">' + $($MatchStatusToServices[$object.Name]) + "</font>" + "</td></tr>" }
            #$mailbody += "<tr><td>" + $object.DisplayName + "</td><td>" + $object.Name + "</td><td>" + '<font color="red">' + $object.Status + '</font>' + "</td><td>" + '<font color="red">' + $($MatchStatusToServices[$object.Name]) + "</font>" + "</td></tr>"
        }#>
        if ($object.Status -eq "Running" -and $MatchStatusToServices[$object.Name] -eq "Running") {
            $mailbody += "<tr><td>" + $object.DisplayName + "</td><td>" + $object.Name + "</td><td>" + '<font color="green">' + $object.Status + '</font>' + "</td><td>" + '<font color="green">' + $($MatchStatusToServices[$object.Name]) + "</font>" + "</td></tr>"
        }
        elseif ($object.Status -eq "Running" -and $MatchStatusToServices[$object.Name] -ne "Running") {
            $mailbody += "<tr><td>" + $object.DisplayName + "</td><td>" + $object.Name + "</td><td>" + '<font color="green">' + $object.Status + '</font>' + "</td><td>" + '<font color="red">' + $($MatchStatusToServices[$object.Name]) + "</font>" + "</td></tr>"
        }
        elseif ($object.Status -ne "Running" -and $MatchStatusToServices[$object.Name] -eq "Running") {
            $mailbody += "<tr><td>" + $object.DisplayName + "</td><td>" + $object.Name + "</td><td>" + '<font color="red">' + $object.Status + '</font>' + "</td><td>" + '<font color="green">' + $($MatchStatusToServices[$object.Name]) + "</font>" + "</td></tr>"
        }
        elseif ($object.Status -ne "Running" -and $MatchStatusToServices[$object.Name] -ne "Running") {
            $mailbody += "<tr><td>" + $object.DisplayName + "</td><td>" + $object.Name + "</td><td>" + '<font color="red">' + $object.Status + '</font>' + "</td><td>" + '<font color="red">' + $($MatchStatusToServices[$object.Name]) + "</font>" + "</td></tr>"
        }
    }
}
else {
    $srxEnv.ResultMessage += "Error in Comparision Services or Service are empty!`n"
    $mailbody += '<font color="red"> Error in Comparision Services or Service are empty - Please check! </font> <br />'
}

$mailbody += '</table></p>'

#Port HTML
$mailbody += '<p>'
$mailbody += '-' * 100 + "<br />"
$mailbody += "Ports Health Check for the Server " + '<u>' + '<b>' + "$InputServer" + '</u>' + '</b>' + " is as below:" + "<br />"
$mailbody += '-' * 100 + "<br />"
$mailbody += '</p>'
$mailbody += '<p><table>'
$mailbody += '<th>ServerName</th><th>Port</th><th>TcpTestSucceeded</th>'

Write-Host "---------------------------------------------------------------"
Write-Host "Please wait we are scanning the necessary exchange ports from server......!"

$Ports = @(  
    [pscustomobject]@{ServerName = '$InputServer'; Port = '25' }  
    [pscustomobject]@{ServerName = '$InputServer'; Port = '80' }
    [pscustomobject]@{ServerName = '$InputServer'; Port = '443' } 
)  
    
foreach ($Port in $Ports) {
    $Port = $Port.Port
    try {
        $Portresult = Test-NetConnection $InputServer -Port $Port
    }
    catch {
        $srxEnv.ResultMessage += "Error to connect and check port $($Port)`n"
        $mailbody += '<font color="red">' + "Error to connect and check port $($Port) - Please check!" + '</font>' + "<br />"
        continue
    }
    
    if ($Portresult.TcpTestSucceeded -eq "True") {
        $mailbody += "<tr><td>" + $PortResult.ComputerName + "</td><td>" + $PortResult.RemotePort + "</td><td>" + '<font color="green">' + $PortResult.TcpTestSucceeded + '</font>' + "</td></tr>"
    }
    else {
        $mailbody += "<tr><td>" + $PortResult.ComputerName + "</td><td>" + $PortResult.RemotePort + "</td><td>" + '<font color="red">' + $PortResult.TcpTestSucceeded + '</font>' + "</td></tr>"
    }   
}

$mailbody += '</table></p>' + "<br />"

Write-Host "---------------------------------------------------------------"
Write-Host "Getting Status for IISAppPool.....!"

#IISAppPoolStatus
$mailbody += '<p>'
$mailbody += '-' * 100 + "<br />"
$mailbody += "IIS Application Pool Status for the Server " + '<u>' + '<b>' + "$InputServer" + '</u>' + '</b>' + " is as below:" + "<br />"
$mailbody += '-' * 100
$mailbody += '</p>'
$mailbody += '<p><table>'
$mailbody += '<th>Name</th><th>Status</th>'

try {
    $AppPool = Invoke-Command -ComputerName $InputServer -ScriptBlock { Get-IISAppPool }
}
catch {
    $srxEnv.ResultMessage += "Error to get App pools`n"
    exit 3
}

if (![string]::IsNullOrEmpty($AppPool)) {
    foreach ($object in $AppPool) {
        if ($object.State -eq 'Started') {
            $mailbody += "<tr><td>" + $object.Name + "</td><td>" + '<font color="green">' + $object.State + '</font>' + "</td></tr>"
        }
        else {
            $mailbody += "<tr><td>" + $object.Name + "</td><td>" + '<font color="red">' + $object.State + '</font>' + "</td></tr>"
        }
    }
}
else {
    $srxEnv.ResultMessage += "Error in AppPool - Please check!`n"
    $mailbody += '<font color="red"> Error in service health status report - Please check! </font> <br />'
}
$mailbody += '</table></p><br />'

#ServerHealthReport
Write-Host "---------------------------------------------------------------"
Write-Host "Getting the overall Health Status Report for the $InputServer!"

$mailbody += '<p>'
$mailbody += '-' * 100 + "<br />"
$mailbody += "Health Report Status for the Server " + '<u>' + '<b>' + "$InputServer" + '</u>' + '</b>' + " is as below:" + "<br />"
$mailbody += '-' * 100
$mailbody += '</p>'
$mailbody += '<p><table>'
$mailbody += '<th>Health Set Name</th><th>Alert Value</th>'
try {
    $ServerHealthStatusReport = Get-HealthReport -Identity $InputServer | Select-Object AlertValue, HealthSet
}
catch {
    $srxEnv.ResultMessage += "Error to get server health status report.`n"
    exit 4
}

if (![string]::IsNullOrEmpty($ServerHealthStatusReport)) {
    foreach ($object in $ServerHealthStatusReport) {
        if ($object.AlertValue -eq "Healthy") {
            $mailbody += "<tr><td>" + $object.HealthSet + "</td><td>" + '<font color="Green">' + $object.AlertValue + '</font></td></tr>'
        }
        elseif ($object.AlertValue -eq "Unhealthy") {
            $mailbody += "<tr><td>" + $object.HealthSet + "</td><td>" + '<font color="Red">' + $object.AlertValue + '</font></td></tr>'
        }
        else {
            $mailbody += "<tr><td>" + $object.HealthSet + "</td><td>" + '<font color="blue">' + $object.AlertValue + '</font></td></tr>'
        } 
    }
}
else {
    $srxEnv.ResultMessage += "Error in service health status report - Please check!`n"
    $mailbody += '<font color="red"> Error in service health status report - Please check! </font> <br />'
}  

$mailbody += '</table></p>' + "<br />"



$mailbody += '-' * 100 + "<br />"
$mailbody += '<font color="red">' + "This is an automatically generated email - please do not reply!" + '</font>' + "<br />"
$mailbody += '-' * 100

$mailbody2 = $mailbody
$mailbody += '</body>'
$mailbody += '</html>'

$mailbody2 += '-' * 100 + "<br />"
$mailbody2 += '<font color="red">' + "This mail is sent over the server $InputServer!" + '</font>' + "<br />"
$mailbody2 += '-' * 100
$mailbody2 += '</body>'
$mailbody2 += '</html>'

try {
    #Mail deliver to selected server
    Send-MailMessage -from $EmailFrom -To $EmailTo -Subject $Subject -body $mailbody -BodyAsHtml -DeliveryNotificationOption OnSuccess, OnFailure -SmtpServer smtpdce.brose.net
}
catch {
    $srxEnv.ResultMessage += "`nFailed to send mail. via Receive Connector"
    exit 5
}

try {
    #Mail deliver to selected server
    Send-MailMessage -from $EmailFrom -To $EmailTo -Subject $Subject -body $mailbody2 -BodyAsHtml -DeliveryNotificationOption OnSuccess, OnFailure -SmtpServer $InputServer
    
}
catch {
    $srxEnv.ResultMessage += "`nFailed to send mail. via $InputServer"
    exit 6
}

Write-Host "--------------------------------------------------------------"
$srxEnv.ResultMessage += "Check your outlook if email is delivered"
$mailbody += '</body>'
$mailbody += '</html>'
