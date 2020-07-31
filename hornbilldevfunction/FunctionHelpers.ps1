function Get-HornbillSession {
    [CmdletBinding()]
    param(
        [string]$ID,
        [string]$Key
    ) 
    Set-HB-Instance -Instance $ID -Key $Key
    return Invoke-HB-XMLMC "session" "getSessionInfo"
}

function Get-MyRequests {
    [CmdletBinding()]
    param(
        [string]$ID,
        [string]$Key,
        [string]$OwnerID,
        [string]$TeamID,
        [string]$TeamName,
        [string]$Direction
    )
    Set-HB-Instance -Instance $ID -Key $Key
    Add-HB-Param "statusEquals" "status.open"
    Add-HB-Param "ownerEquals" $OwnerID
    Add-HB-Param "teamEquals" $TeamID
    Add-HB-Param "rowstart" "0"
    Add-HB-Param "limit" "3"
    Add-HB-Param "orderByColumn" "h_datelogged"
    Add-HB-Param "orderByDirection" $Direction
    Add-HB-Param "visibleColumns" "referenceColumn,customerColumn,companyColumn,serviceColumn,catalogColumn,summaryColumn,createdByColumn,raisedOnColumn,priorityColumn"
    $Requests = Invoke-HB-XMLMC "apps/com.hornbill.servicemanager/Requests" "smGetRequests"
    $RequestRows = $Requests.params.requests | ConvertFrom-Json
    if ($RequestRows.row.Count -eq 0) {
        $MessageBody = "You own no Requests in <strong>$($TeamName)</strong>"
    } else{
        $MessageBody = "The $($RequestRows.row.Count) "
        if ($Direction -eq "descending") {
            $MessageBody += "most recently logged"
        } else {
            $MessageBody += "oldest"
        }
        $MessageBody += " requests that are assigned to you in <strong>$($TeamName):</strong><hr>"
    }
    foreach ($RD in $RequestRows.row) {
        $MessageBody += "<h1><a href=`"https://live.hornbill.com/$($InstanceName)/servicemanager/request/view/$($RD.h_pk_reference)`">$($RD.h_requesttype) - $($RD.h_pk_reference)</a></h1>"
        $MessageBody += "<p>$($RD.h_summary)</p>"
        $ReqStatus = (Get-Culture).textinfo.totitlecase($RD.h_status.Split('.')[1])
        $MessageBody += "<p><strong>Status</strong>: $($ReqStatus)</p>"
        $MessageBody += "<p><strong>Customer</strong>: $($RD.customer_name)</p>"
        $MessageBody += "<p><strong>Logged At</strong>: $($RD.h_datelogged)</p>"
        if ($null -ne $RD.h_service_name) {
            $MessageBody += "<p><strong>Service</strong>: $($RD.h_service_name)"
            if ($null -ne $RD.h_catalog ) {
                $MessageBody += " - $($RD.h_catalog)"
            }
            $MessageBody += "</p>"
        }
        if ($null -ne $RD.h_fk_priorityname) {
            $MessageBody += "<p><strong>Priority</strong>: $($RD.h_fk_priorityname)</p>"
        }
        $MessageBody +="<hr>"
    }     
    $bodyObj = @{
        "type" = "message"
        "text" = $MessageBody
    }
    return $bodyObj
}

function Get-TeamRequests {
    [CmdletBinding()]
    param(
        [string]$ID,
        [string]$Key,
        [string]$UserID,
        [string]$TeamID,
        [string]$TeamName,
        [string]$Direction
    )
    Set-HB-Instance -Instance $ID -Key $Key
    Add-HB-Param "statusEquals" "status.open"
    Add-HB-Param "ownerNotEquals" $UserID
    Add-HB-Param "teamEquals" $TeamID
    Add-HB-Param "rowstart" "0"
    Add-HB-Param "limit" "3"
    Add-HB-Param "orderByColumn" "h_datelogged"
    Add-HB-Param "orderByDirection" $Direction
    Add-HB-Param "visibleColumns" "referenceColumn,customerColumn,companyColumn,serviceColumn,catalogColumn,summaryColumn,createdByColumn,raisedOnColumn,ownerColumn,priorityColumn"
    $Requests = Invoke-HB-XMLMC "apps/com.hornbill.servicemanager/Requests" "smGetRequests"
    $RequestRows = $Requests.params.requests | ConvertFrom-Json
    if ($RequestRows.row.Count -eq 0) {
        $MessageBody = "There are no requests that are not assigned to you and owned by <strong>$($TeamName)</strong>."
    } else{
        $MessageBody = "The $($RequestRows.row.Count) "
        if ($Direction -eq "descending") {
            $MessageBody += "most recently logged"
        } else {
            $MessageBody += "oldest"
        }
        $MessageBody += " requests not assigned to you and owned by <strong>$($TeamName)</strong>:<hr>"
    }

    foreach ($RD in $RequestRows.row) {
        $MessageBody += "<h1><a href=`"https://live.hornbill.com/$($InstanceName)/servicemanager/request/view/$($RD.h_pk_reference)`">$($RD.h_requesttype) - $($RD.h_pk_reference)</a></h1>"
        $MessageBody += "<p>$($RD.h_summary)</p>"
        $ReqStatus = (Get-Culture).textinfo.totitlecase($RD.h_status.Split('.')[1])
        $MessageBody += "<p><strong>Status</strong>: $($ReqStatus)</p>"
        $MessageBody += "<p><strong>Customer</strong>: $($RD.customer_name)</p>"
        $MessageBody += "<p><strong>Logged At</strong>: $($RD.h_datelogged)</p>"
        if ($null -ne $RD.h_service_name) {
            $MessageBody += "<p><strong>Service</strong>: $($RD.h_service_name)"
            if ($null -ne $RD.h_catalog ) {
                $MessageBody += " - $($RD.h_catalog)"
            }
            $MessageBody += "</p>"
        }
        if ($null -ne $RD.h_ownername) {
            $MessageBody += "<p><strong>Owner</strong>: $($RD.h_ownername)</p>"
        } else {
            $MessageBody += "<p><strong>Owner</strong>: No Owner</p>"
        }
        if ($null -ne $RD.h_fk_priorityname) {
            $MessageBody += "<p><strong>Priority</strong>: $($RD.h_fk_priorityname)</p>"
        }
        $MessageBody +="<hr>"
    }     
    $bodyObj = @{
        "type" = "message"
        "text" = $MessageBody
    }
    return $bodyObj
}

function Get-RequestDetails {
    [CmdletBinding()]
    param(
        [string]$ID,
        [string]$Key,
        [string]$RequestRef
    )
    Set-HB-Instance -Instance $ID -Key $Key
    Add-HB-Param "application" "com.hornbill.servicemanager"
    Add-HB-Param "entity" "Requests"
    Add-HB-Param "keyValue" $RequestRef
    Add-HB-Param "formatValues" "true"
    $Request = Invoke-HB-XMLMC "data" "entityGetRecord"
    return $Request
}

function Get-Request {
    [CmdletBinding()]
    param(
        [string]$ID,
        [string]$Key,
        [string]$RequestRef,
        [string]$TeamID,
        [string]$TeamName
    )
    $Request = Get-RequestDetails $ID $Key $RequestRef
    $MessageBody = ""
    
    if($null -eq $Request.Params.primaryEntityData) {
        $MessageBody += "Request $($RequestRef) not found."      
    } else {
        $RD = $Request.Params.primaryEntityData.record
        if ($RD.h_fk_team_id -ne $TeamID) {
            $MessageBody += "Request $($RequestRef) isn't owned by $($TeamName), so can't be viewed in this Team."      
        } else {
            $MessageBody += "<h1><a href=`"https://live.hornbill.com/$($InstanceName)/servicemanager/request/view/$($RD.h_pk_reference)`">$($RD.h_requesttype) - $($RD.h_pk_reference)</a></h1>"
            $MessageBody += "<p>$($RD.h_summary)</p>"
            $ReqStatus = (Get-Culture).textinfo.totitlecase($RD.h_status.Split('.')[1])
            $MessageBody += "<p><strong>Status</strong>: $($ReqStatus)</p>"
            $MessageBody += "<p><strong>Customer</strong>: $($RD.h_fk_user_name)</p>"
            $MessageBody += "<p><strong>Logged At</strong>: $($RD.h_datelogged)</p>"
            if ($null -ne $RD.h_fk_servicename) {
                $MessageBody += "<p><strong>Service</strong>: $($RD.h_fk_servicename)"
                if ($null -ne $RD.h_catalog ) {
                    $MessageBody += " - $($RD.h_catalog)"
                }
                $MessageBody += "</p>"
            }
            if ($null -ne $RD.h_ownername) {
                $MessageBody += "<p><strong>Owner</strong>: $($RD.h_ownername)</p>"
            } else {
                $MessageBody += "<p><strong>Owner</strong>: No Owner</p>"
            }
            if ($null -ne $RD.h_fk_priorityname) {
                $MessageBody += "<p><strong>Priority</strong>: $($RD.h_fk_priorityname)</p>"
            }
            $MessageBody += "<p><strong>Description</strong>: $($RD.h_description)</p>"
        }
    }
    $bodyObj = @{
        "type" = "message"
        "text" = $MessageBody
    }
    return $bodyObj
}
function Update-TakeRequest {
    [CmdletBinding()]
    param(
        [string]$ID,
        [string]$Key,
        [string]$OwnerID,
        [string]$OwnerName,
        [string]$TeamID,
        [string]$TeamName,
        [string]$RequestRef
    )
    $Request = Get-RequestDetails $ID $Key $RequestRef
    $MessageBody = ""
    
    if($null -eq $Request.Params.primaryEntityData) {
        $MessageBody += "Request $($RequestRef) not found."      
    } else {
        $RD = $Request.Params.primaryEntityData.record
        if ($RD.h_fk_team_id -ne $TeamID) {
            $MessageBody += "Request $($RequestRef) isn't owned by $($TeamName), so can't be taken by a user in this Team."      
        } else {
            Set-HB-Instance -Instance $ID -Key $Key
            Add-HB-Param "inReference" $RequestRef
            Add-HB-Param "inAssignToId" $OwnerID
            Add-HB-Param "inAssignToGroupId" $TeamID
            
            $Timeline = @{
                "updateText" = "''Assigned by $($OwnerName) via Microsoft Teams''"
                "source" = "teams"
                "extra" = @{
                    "h_ownername" = $OwnerName
                    "h_fk_team_id" = $TeamID
                    "h_group_name" = $TeamName
                }
                "visibility" = "trustedGuest"
            }
            $TimelineJSON = $Timeline | ConvertTo-JSON
            Add-HB-Param "updateTimelineInputs" $TimelineJSON
            $APICall = Invoke-HB-XMLMC "apps/com.hornbill.servicemanager/Requests" "assign"
            if ($APICall.Status -eq "ok") {
                if ($null -ne $APICall.Params.assignError -and $APICall.Params.assignError -ne "") {
                    $MessageBody = "An error occurred when attempting to assign the request."
                } else {
                    $MessageBody = "$($RequestRef) has been assigned to you in $($TeamName)"
                }
            } else{
                $MessageBody = "An error occurred when attempting to assign the request."
            }
        }
    }    
    $bodyObj = @{
        "type" = "message"
        "text" = $MessageBody
    }
    return $bodyObj
}

function Update-Request {
    [CmdletBinding()]
    param(
        [string]$ID,
        [string]$Key,
        [string]$UpdateBy,
        [string]$RequestRef,
        [string]$Update,
        [string]$TeamID,
        [string]$TeamName
    )
    $Request = Get-RequestDetails $ID $Key $RequestRef
    $MessageBody = ""
    
    if($null -eq $Request.Params.primaryEntityData) {
        $MessageBody += "Request $($RequestRef) not found."      
    } else {
        $RD = $Request.Params.primaryEntityData.record
        if ($RD.h_fk_team_id -ne $TeamID) {
            $MessageBody += "Request $($RequestRef) isn't owned by $($TeamName), so can't be taken by a user in this Team."      
        } else {
            Set-HB-Instance -Instance $ID -Key $Key
            Add-HB-Param "requestId" $RequestRef
            Add-HB-Param "source" "Microsoft Teams"
            Add-HB-Param "content" "''Update provided by $($UpdateBy) via Microsoft Teams''`n`n$($Update)"
            Add-HB-Param "visibility" "trustedGuest"
            Add-HB-Param "activityType" "Update"
            $APICall = Invoke-HB-XMLMC "apps/com.hornbill.servicemanager/Requests" "updateReqTimeline"
            if ($APICall.Status -eq "ok") {
                if ($null -ne $APICall.Params.exceptionDescription -and $APICall.Params.exceptionDescription -ne "") {
                    $MessageBody = "An error occurred when attempting to assign the request."
                } else {
                    $MessageBody = "$($RequestRef) has been updated successfully."
                }
            } else{
                $MessageBody = "An error occurred when attempting to assign the request."
            }
        }
    }    
    $bodyObj = @{
        "type" = "message"
        "text" = $MessageBody
    }
    return $bodyObj
}

function Get-HornbillTeam {
    [CmdletBinding()]
    param(
        [string]$ID,
        [string]$Key,
        [string]$TeamID
    )
    Set-HB-Instance -Instance $ID -Key $Key
    Add-HB-Param "id" $TeamID
    $GroupDetails = Invoke-HB-XMLMC "admin" "groupGetInfo"
    return $GroupDetails 
}

function Get-AzureUser {
    [CmdletBinding()]
    param(
        [string]$ID,
        [string]$Key,
        [int]$KeySafe,
        [string]$UserID
    )

    $Payload = @{
        "UPN" = $UserID
    }
    $JSONPayload = $Payload | ConvertTo-JSON -Compress

    Set-HB-Instance -Instance $ID -Key $Key
    Add-HB-Param "methodPath" "/Microsoft/Azure/Users/Get User.m"
    Add-HB-Param "requestPayload" $JSONPayload
    Open-HB-Element "credential"
    Add-HB-Param "id" "microsoft"
    Add-HB-Param "keyId" $KeySafe.ToString()
    Close-HB-Element "credential"
    $UserDetails = Invoke-HB-XMLMC "bpm" "iBridgeInvoke"
    return $UserDetails
}