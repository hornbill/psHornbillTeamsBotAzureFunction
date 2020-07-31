using namespace System.Net
# Input bindings are passed in via param block.
param($Request)

# Hornbill Instance Details
$InstanceName = "instanceid"
$InstanceKey = "apikey"
$InstanceMicrosoftKey = 1

# Map your Microsoft Teams against Hornbill Teams, providing Custom Outgoing Webhook Security Tokens for each to authorise requests
$TeamsMapping = @{
    "19:TEAMSTEAMID1@thread.skype" = @{
        "HornbillTeamID" = "2ndLineSupport"
        "WebhookToken" = "WEBHOOK1SECURITYTOKEN"
    }
    "19:TEAMSTEAMID2@thread.tacv2" = @{
        "HornbillTeamID" = "3rdLineSupport"
        "WebhookToken" = "WEBHOOK2SECURITYTOKEN"
    }
}

# Authorize requests with HMACSHA256 
$secret = $TeamsMapping[$Request.Body.channelData.teamsTeamId]["WebhookToken"]
$hmacsha = New-Object System.Security.Cryptography.HMACSHA256
$hmacsha.key = [Convert]::FromBase64String($secret)
$signature = $hmacsha.ComputeHash([Text.Encoding]::UTF8.GetBytes($Request.RawBody))
$signature = [Convert]::ToBase64String($signature)
if ($Request.Headers.authorization -ne "HMAC $($signature)") {
    # Authorization Header doesn't 
    $bodyObj = @{type="message";text=""} 
    $bodyObj.text = "This request is not authorized. Speak to your "
    $body = $bodyObj | ConvertTo-Json
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::OK
        Body = $body
    })
    Exit 0
}

. "./hornbilldevfunction/FunctionHelpers.ps1"

$status = [HttpStatusCode]::OK

$TeamDetails = Get-HornbillTeam $InstanceName $InstanceKey $TeamsMapping[$Request.Body.channelData.teamsTeamId]["HornbillTeamID"]
$UserDetails = Get-AzureUser $InstanceName $InstanceKey $InstanceMicrosoftKey $Request.Body.from.aadObjectId
$User = ($UserDetails.Params.responsePayload | ConvertFrom-JSON).params

if ($Request.Body.text -like '*my requests old*' -or $Request.Body.text -like '*my old requests*' -or $Request.Body.text -like '*old my requests*') {
    $bodyObj = Get-MyRequests $InstanceName $InstanceKey $User.userPrincipalName $TeamDetails.Params.id $TeamDetails.Params.name "ascending"
    $body = $bodyObj | ConvertTo-Json -Depth 4
} elseif ($Request.Body.text -like '*my requests*' -or $Request.Body.text -like '*my new requests*'  -or $Request.Body.text -like '*my requests new*') {
    $bodyObj = Get-MyRequests $InstanceName $InstanceKey $User.userPrincipalName $TeamDetails.Params.id $TeamDetails.Params.name "descending"
    $body = $bodyObj | ConvertTo-Json -Depth 4
} elseif ($Request.Body.text -like '*team requests old*' -or $Request.Body.text -like '*team old requests*' -or $Request.Body.text -like '*old team requests*') {
    $bodyObj = Get-TeamRequests $InstanceName $InstanceKey $User.userPrincipalName $TeamDetails.Params.id $TeamDetails.Params.name "ascending"
    $body = $bodyObj | ConvertTo-Json -Depth 4
} elseif ($Request.Body.text -like '*team requests*' -or $Request.Body.text -like '*team new requests*' -or $Request.Body.text -like '*team requests new*') {
    $bodyObj = Get-TeamRequests $InstanceName $InstanceKey $User.userPrincipalName $TeamDetails.Params.id $TeamDetails.Params.name "descending"
    $body = $bodyObj | ConvertTo-Json -Depth 4
} elseif ($Request.Body.text -like '*get request*') {
    $StringArray = $Request.Body.text.Split(" ")
    $bodyObj = Get-Request $InstanceName $InstanceKey $StringArray[$StringArray.Count-1] $TeamDetails.Params.id $TeamDetails.Params.name
    $body = $bodyObj | ConvertTo-Json -Depth 4
} elseif ($Request.Body.text -like '*take request*') {
    $StringArray = $Request.Body.text.Split(" ")
    $bodyObj = Update-TakeRequest $InstanceName $InstanceKey $User.userPrincipalName $User.displayName $TeamDetails.Params.id $TeamDetails.Params.name $StringArray[$StringArray.Count-1]
    $body = $bodyObj | ConvertTo-Json -Depth 4
} elseif ($Request.Body.text -like '*update request*') {
    $StringArray = $Request.Body.text -split "update request",2
    $ContentString = $StringArray[$StringArray.Count - 1].TrimStart()
    $ContentArray = $ContentString -split " ",2
    $bodyObj = Update-Request $InstanceName $InstanceKey $User.displayName $ContentArray[0] $ContentArray[1] $TeamDetails.Params.id $TeamDetails.Params.name
    $body = $bodyObj | ConvertTo-Json -Depth 4
} else {
    $MessageBody = "I don't understand 🙁 You can ask me the following:"
    if ($Request.Body.text -like '*help*') {
        $MessageBody = "Sure! You can ask me any of the following:"
    }
    $MessageBody += "<ul>"
    $MessageBody += "<li><strong>my requests</strong>: will return details of the latest 3 requests logged that are owned by you in this group</li>"
    $MessageBody += "<li><strong>my requests old</strong>: will return details of the oldest 3 requests logged that are owned by you in this group</li>"
    $MessageBody += "<li><strong>team requests</strong>: will return details of the latest 3 requests logged that are not owned by you, but are assigned to this group</li>"
    $MessageBody += "<li><strong>team requests old</strong>: will return details of the oldest 3 requests logged that are not owned by you, but are assigned to this group</li>"
    $MessageBody += "<li><strong>take request IN0000123</strong>: will assign request <strong>IN0000123</strong> to you in this group</li>"
    $MessageBody += "<li><strong>update request IN0000123 some update text</strong>: will add an update to the timeline of request <strong>IN0000123</strong> with <strong>some update text</strong</li>"
    $MessageBody += "</ul>"
    $bodyObj = @{type="message";text=$MessageBody} 
    $body = $bodyObj | ConvertTo-Json
}
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
    StatusCode = $status
    Body = $body
})
