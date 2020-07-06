# Set up file and path variables
$resultsDir = "%BA_DRIVE%:\runtime-data\agent-data\work\%CHECKOUT_FOLDER%"
$inputFile = "$resultsDir\SpecRunTestResults_FrontEnd.html"
$zipPath = "$resultsDir\archive.zip"
$specRunLog = "E:\runtime-data\agent-data\work\Hitched-Uat\Hitched.InterfaceTests.FrontEnd\bin\specrun.log"

# Read in results file
$content = [IO.File]::ReadAllText($inputFile)

# Build agent information
$psVerInfo = get-host
$dotNetVerInfo = [Runtime.InteropServices.RuntimeEnvironment]::GetRuntimeDirectory()

# Zip up results
if(-not (test-path($zipPath)))
{
	set-content $zipPath ("PK" + [char]5 + [char]6 + ("$([char]0)" * 18))
	(dir $zipPath).IsReadOnly = $false	
}
$shellApplication = new-object -com shell.application
$zipPackage = $shellApplication.NameSpace($zipPath)
$zipPackage.CopyHere($inputFile)
Start-sleep -s 30
$zipPackage.CopyHere($specRunLog)
Start-sleep -s 30

# Function to find inner text between two tags (i.e. strip tags)
function Find-InnerText {
    param( [string] $tagStart, [string] $tagEnd, [string] $textToSearch )
    $indexOfTagStart = $textToSearch.IndexOf($tagStart)
    $indexOfTagEnd = $textToSearch.IndexOf($tagEnd,$indexOfTagStart)
    $indexOfInnerTextStart = $indexOfTagStart + $tagStart.Length
    $lengthOfInnerText = $indexOfTagEnd - $indexOfTagStart - $tagStart.length
    $textToSearch.Substring($indexOfInnerTextStart,$lengthOfInnerText).Trim()
}

# Pull out HTML for email

# Determine pass rate
$FileExists = Test-Path $inputFile
If ($FileExists -eq $True) {$PassRate = Find-InnerText "<td id = `"PassRate`">" "</td>" $content}
Else {$PassRate = "No recorded result"}

# Get CSS
$StyleContent = Find-InnerText "<style type=`"text/css`">" "</style>" $content 
$Style = "<style type=`"text/css`">" + $StyleContent + "</style>"

# Get Summary
$SummaryContent = Find-InnerText "<h1>Hitched.InterfaceTests.FrontEnd Test Execution Report</h1>" "</ul>" $content
$Summary = "<h1>Hitched.InterfaceTests.FrontEnd Test Execution Report</h1>" + $SummaryContent + "</ul>"

# Get Times
$StartTime = Find-InnerText "<li>Start Time:" "</li>" $SummaryContent
$Duration = Find-InnerText "<li>Duration:" "</li>" $SummaryContent

# Get Results Table
$FirstResultsTableContent = Find-InnerText "<table class=`"testEvents`">" "<h2>Test Timeline Summary</h2>" $content
$FirstResultsTable = "<table class=`"testEvents`">" + $FirstResultsTableContent 

# Get Stats
$Tests = Find-InnerText "<td id = `"Tests`">" "</td>" $FirstResultsTableContent
$Succeeded = Find-InnerText "<td id = `"Succeeded`">" "</td>" $FirstResultsTableContent
$Failed = Find-InnerText "<td id = `"Failed`">" "</td>" $FirstResultsTableContent
$Ignored = Find-InnerText "<td id = `"Ignored`">" "</td>" $FirstResultsTableContent

# Get Summary Content
$FeatureSummaryContent = Find-InnerText "<h2>Feature Summary</h2>" "<h2>Error Summary</h2>" $content
$FeatureSummary = "<h2>Feature Summary</h2>" + $FeatureSummaryContent

# Start of email
$Start = "
<p>Full results attached or 
<a href=`"<teamcityreportlink>`">
 Click Here</a> (Team City Credentials Required)</p>
"

# Construct email
$From = "automated@test.results"
$To = "<my email address>"
$Attachment = "$zipPath"

$BuildAgentInfo = "<strong> Build Agent Info </strong><br>" + "Powershell Version: " + $PSVersionTable.PSVersion + "<br>" + ".net Version: " + $dotNetVerInfo + "<br>"
$Email = "<!DOCTYPE html><html><body>" + $Start + "<br>" + $Summary + "<br>" + $FirstResultsTable + "<br>" + $FeatureSummaryContent + "<br>" + $BuildAgentInfo + "</body></html>"

$smtpServer = "<Ip Address>"

$message = New-Object System.Net.Mail.MailMessage $From, $To
$message.Subject = "Hitched SpecRunTestResult - $PassRate"
$message.IsBodyHTML = $true
$message.Body = ConvertTo-Html -Body $Email -Head $Style
$message.Attachments.Add($Attachment)
$message.To.Add("<other email address")

# Send email
$smtp = New-Object Net.Mail.SmtpClient($smtpServer)
$smtp.Send($message)

# Slack end point for hitched-imt channel
$uriSlack = "<slack web hook>"

# body for request
$body = @"
{
    "pretext": "Test Results",
    "text": "Message_PH",
    "type": "mrkdwn",
    "color": "#775a97"
}
"@

$MessageText = 
    "*Pass Rate* : " + $PassRate + "\n" +
    "*Start* : " + $StartTime + "\n" +
    "*Duration* : " + $Duration + "\n" +
    "*Tests* : " + $Tests + "\n" +
    "*Succeeded* : " + $Succeeded + "\n" +
    "*Failed* : " + $Failed + "\n" +
    "*Ignored* : " + $Ignored + "\n"

$body = $body -replace "Message_PH", $MessageText

# Make Slack Request
try 
{
    $request = [System.Net.WebRequest]::Create($uriSlack)
    $request.ContentType = "application/json"
    $request.Method = "POST"

    try
    {
        $requestStream = $request.GetRequestStream()
        $streamWriter = New-Object System.IO.StreamWriter($requestStream)
        $streamWriter.Write($body)
    }

    catch
    {
        $BodyErrorMessage = $_.Exception.Message
    }

    finally
    {
        if ($null -ne $streamWriter) { $streamWriter.Dispose() }
        if ($null -ne $requestStream) { $requestStream.Dispose() }
    }

    # Force TLS 1.2
    [Net.ServicePointManager]::SecurityProtocol =  [Enum]::ToObject([Net.SecurityProtocolType], 3072)
    $res = $request.GetResponse()

} 
catch 
{
    $_.Exception.Message
}

# File upload - if curl worked on powershell 2.0 could use this
# $uploadFileComment = "initial_comment=Test Results Report"
# $uploadFileAuth = "Authorization: Bearer <Slack Auth Token>"
# $uploadFileChannel = "channels=<Channel ID>"
# $uploadFileUrl = "https://slack.com/api/files.upload"

# curl -F $Attachment -F $uploadFileComment -F $uploadFileChannel -H $uploadFileAuth $uploadFileUrl

# Slack end point for file upload
$uploadFileUrl = "https://slack.com/api/files.upload"

# Results file
# Build webclient for request
$webclient = new-object system.net.webclient
$webClient.Headers.Add("Content-Type", "application/x-www-form-urlencoded")

# Build parameters as query string
$NVC = New-Object System.Collections.Specialized.NameValueCollection
$NVC.Add("content", $content)
$NVC.Add("filetype", "html")
$NVC.Add("title", "TestResults")
$NVC.Add("initial_comment", "Test Results")
$NVC.Add("token", "<Slack Auth Token>")
$NVC.Add("channels", "<Channel ID>")

# Try request 
try 
{
    $result = $webclient.UploadValues($uploadFileUrl, "POST", $NVC)
    # Uncomment below to see response
    # [System.Text.Encoding]::UTF8.GetString($Result)
}
catch 
{
    $_.Exception.Message
}

# Specrun log
$logContent = [IO.File]::ReadAllText($specRunLog)

# Build webclient for request
$newWebClient = new-object system.net.webclient
$newWebClient.Headers.Add("Content-Type", "application/x-www-form-urlencoded")

# Build parameters as query string
$newNVC = New-Object System.Collections.Specialized.NameValueCollection
$newNVC.Add("content", $logContent)
$newNVC.Add("filetype", "log")
$newNVC.Add("title", "SpecrunLog")
$newNVC.Add("initial_comment", "SpecrunLog")
$NVC.Add("token", "<Slack Auth Token>")
$NVC.Add("channels", "<Channel ID>")

# Try request 
try 
{
    $newResult = $newWebclient.UploadValues($uploadFileUrl, "POST", $newNVC)
    # Uncomment below to see response
    # [System.Text.Encoding]::UTF8.GetString($newResult)
}
catch 
{
    $_.Exception.Message
}
