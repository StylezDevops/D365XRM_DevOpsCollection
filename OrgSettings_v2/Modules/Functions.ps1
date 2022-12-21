# XRM functions
function InstallAndImport {
    param 
    ( 
        [Parameter(Mandatory = $true, HelpMessage = "Array of Powershell Modules")] [Object]$PSModules
    )
    # InstallAndImport -PSModules "SomeModule", "SomeOtherModule", "AnotherModule"
    try {
        # Import required module(s)
        $PSModules | Foreach-Object {
            Write-Output "Importing $_"
            Install-Module -Name $_ -AllowClobber -Force -ErrorAction Stop -Scope CurrentUser
            Import-Module -Name $_ -ErrorAction Stop
            Write-Output "$_ imported successfully."
        }
    }
    catch {
        Write-Output $_.Exception.Message
    }
}
function XrmConnect {
    param 
    ( 
        [Parameter(Mandatory = $true, HelpMessage = "clientId")] [String]$clientId,
        [Parameter(Mandatory = $true, HelpMessage = "clientSecret")] [String]$clientSecret,
        [Parameter(Mandatory = $true, HelpMessage = "crmURL")] [String]$crmURL,
        [Parameter(Mandatory = $false, HelpMessage = "User GuID to impersonate")][String]$impersonation_Guid
    )
    # I did this because I was fedup of passing a second command to have to impersonate.
    # this modification means i can pass impersonation_guid or not when connecting.
    # XrmConnect -clientId $clientId -clientSecret $clientSecret -crmURL $crmURL -impersonation_Guid $impersonation_Guid #$impersonation_Guid is optional
    Write-Output "Connecting to CRM"
    # Force TLS 1.2
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
    $conn = Connect-CrmOnline -OAuthclientId $clientId  -clientSecret $clientSecret -serverUrl $crmURL
    $j = ($conn | Select -ExpandProperty IsReady -ErrorAction SilentlyContinue )
    try {
        if ($j -eq $true) { 
            Write-Output "Connected To Environment"
            if (![string]::IsNullOrEmpty($impersonation_Guid)) { 
                Write-Output "Impersonation GuID provided. Setting CallerID Property to $impersonation_Guid `n"
                Set-CrmConnectionCallerId -CallerId $impersonation_Guid
            }
        }
        elseif ( $j -eq $false ) { throw  "Not Connected to CRM" } elseif ( [string]::IsNullOrEmpty($j) ) { throw "Not Connected to CRM" }
    }
    catch {
        Write-Output ($error[0].Exception.GetType()).Name $error[0].Exception.Message
    }
    return $conn 
}
function WhatIsThePrimaryKey() {
    param(        
        [Parameter(Mandatory = $true, HelpMessage = "Entity Name")] [object]$entityName
    )
    # Microsoft completely ballsed up the naming of the primary key
    # and appear to have forgotten the naming convention at 3 different points in time
    # - when they were developing "activities", "usersettings" and "systemforms"
    # the activities one kind of makes sense. The other 2 are just stupid.
    #
    # WhatIsThePrimaryKey -entityName someEntity
    $OOBActivities = @(
        "opportunityclose",
        "socialactivity",
        "campaignresponse",
        "letter", "orderclose",
        "appointment",
        "recurringappointmentmaster",
        "fax",
        "email",
        "activitypointer",
        "incidentresolution",
        "bulkoperation",
        "quoteclose",
        "task",
        "campaignactivity",
        "serviceappointment",
        "phonecall"
    )
    # set the primaryKey where MS fucked it up
    if ($entityName -eq "usersettings") {
        $primaryKey = "systemuserid"
    }
    elseif ($entityName -eq "systemform") {
        $primaryKey = "formid"
    }
    elseif ($entityName -in $OOBActivities) {
        $primaryKey = "activityid"
    }
    else {
        # default - the regular, sane primaryKey naming convention
        $primaryKey = $entityName + "id"
    }
    
    $primaryKey
}

function GetAccessTokenXRMAPI {
    param
    ( 
        [Parameter(Mandatory = $true, HelpMessage = "clientId")] [String]$clientId,
        [Parameter(Mandatory = $true, HelpMessage = "clientSecret")] [String]$clientSecret,
        [Parameter(Mandatory = $true, HelpMessage = "crmURL")] [String]$crmURL,
        [Parameter(Mandatory = $true, HelpMessage = "Tenant ID")] [String]$tenantId
    )
    # $newToken = XRM-APIgetAccessToken -ClientId $clientId -ClientSecret $clientSecret -crmURL $crmURL -tenantId $tenantId
    # sometimes can't be arsed messing about and this is just easier to connect. Also - I can easily get stuff in json format
    # like this direct from the CRM and use Invoke-RestMthod to pull it right into a pscustomobject
    $oAuthTokenEndpoint = 'https://login.microsoftonline.com/' + $tenantId + '/oauth2/v2.0/token'
    $authBody = 
    @{
        client_id     = $clientId;
        client_secret = $clientSecret;    
        scope         = "$crmURL/.default"    
        grant_type    = 'client_credentials'
    }
    # Token Request Parameters
    $authParams = 
    @{
        URI         = $oAuthTokenEndpoint
        Method      = 'POST'
        ContentType = 'application/x-www-form-urlencoded'
        Body        = $authBody
    }
    # Get Token
    try {
        $authRequest = Invoke-RestMethod @authParams -ErrorAction Stop
        $authResponse = $authRequest
    }
    catch {
        Write-Output ($error[0].Exception.GetType()).Name $error[0].Exception.Message
    }
    return $authResponse
}
function queryEntityXRMAPI {
    param 
    ( 
        [Parameter(Mandatory = $true, HelpMessage = "crmURL")] [String]$crmURL,
        [Parameter(Mandatory = $true, HelpMessage = "entityName")] [String]$entityName,
        [Parameter(Mandatory = $true, HelpMessage = "Auth Resopnse")] [Object]$authResponse
    )
    # query entities once connected to WebAPI with authToken
    # example: $orgObject = (XRM-APIqueryEntity -crmURL $crmURL -entityName "organization" -authResponse $newToken).value
    # and pull the results right back into a PSCustomObject (invoke-restmethod)
    if ($entityName -eq "organization") {
        $entityName = "organizations"
    }
    $apiCallParams =
    @{
        URI     = $crmURL + "/api/data/v9.2/$entityName"
        Headers = @{
            "Authorization"    = "$($authResponse.token_type) $($authResponse.access_token)";
            "Accept"           = "application/json;odata=nometadata";
            "OData-MaxVersion" = "4.0";  
            "OData-Version"    = "4.0";  
  
        }
        Method  = 'GET'
    }
    $apiCallRequest = Invoke-RestMethod @apiCallParams -ErrorAction Stop
    return $apiCallRequest
}
function GetRecordIdWithFilter {
    param 
    ( 
        [Parameter(Mandatory = $true, HelpMessage = "entityName")] [String]$entityName,
        [Parameter(Mandatory = $true, HelpMessage = "Filter Attribute")] [String]$Attribute,
        [Parameter(Mandatory = $true, HelpMessage = "Filter Operator")] [String]$operator,
        [Parameter(Mandatory = $true, HelpMessage = "Filter Value")] [String]$value
    )
    # Returns a recordID
    #ie: to get the ID of the record where the landlordName attribute is "Fred" in the pub entity:
    # GetRecordIdWithFilter -entityName pub -Attribute Name -operator eq -value "Fred"
    # this will return JUST the id of the record (if it exists) -  
    # useful for doing a quicklookup for the id to pass to 'Set-CrmRecord'    
    $entid = (WhatIsThePrimaryKey -entityName $entityName)

    if (![string]::IsNullOrEmpty($Attribute)) {
        try {
            $recordID = (Get-CrmRecords -EntityLogicalName $entityName -FilterAttribute $Attribute -FilterOperator $operator -FilterValue "$value" -Fields "*").CrmRecords.$entid.Guid
        }
        catch {
            Write-Output ($error[0].Exception.GetType()).Name $error[0].Exception.Message
        }
    }
    return $recordID
}
