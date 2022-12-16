# Drop this as a task in a "deploy" job
# and it SHOULD find the artifact uploaded by the other script
# I might need to mess around with paths to get this 
# fully working in a pipeline context..
# if you get a "systemInt32" type error when setting a field..
# add it to the org_picklistfields list. I occasionally find one I missed.
param 
( 
    [Parameter(Mandatory = $true, HelpMessage = "ClientId")] [String]$ClientId,
    [Parameter(Mandatory = $true, HelpMessage = "ClientSecret")] [String]$ClientSecret,
    [Parameter(Mandatory = $true, HelpMessage = "crmURL")] [String]$crmURL,
    [Parameter(Mandatory = $true, HelpMessage = "User GuID to impersonate")][String]$impersonation_Guid
)
. $PSScriptRoot\Modules\Functions.ps1
$orgSettingsJSON = (Get-ChildItem .\jsonExport\).Name
# import modules
InstallAndImport -PSModules "Microsoft.Xrm.Tooling.CrmConnector.PowerShell", "Microsoft.Xrm.Data.PowerShell"
# connect
XrmConnect -ClientId $ClientId -ClientSecret $ClientSecret -crmURL $crmURL -impersonation_Guid $impersonation_Guid #$impersonation_Guid is needed here
# get the json
$json = Get-Content .\jsonExport\$orgSettingsJSON -Encoding UTF8 | Out-String | ConvertFrom-Json
$picklistFields = Get-Content ".\dataTypes\org_picklistfields"
$dateTimeFields = Get-Content ".\dataTypes\org_dateTimeFields"
$orgid = (Invoke-CrmWhoAmI).OrganizationId
# Convert PSCustomObject to hashtable
$hashtable = @{}
#$infohash = @{}
ForEach ($field in $json.PSObject.Properties) {
    $txtStr = $field.Name
    $value = $field.Value
    if ($picklistFields.Contains("$txtStr")) {
        #$infohash[$txtStr] = $field.Value
        $optionSetValue = New-CrmOptionSetValue $field.Value
        $hashtable[$txtStr] = $optionSetValue
    }
    elseif ($dateTimeFields.Contains("$txtStr")) {
        $hashtable[$field.Name] = [datetime]$field.Value 
    }
    else {
        $hashtable[$field.Name] = $field.Value 
    }
}
$hashtable.GetEnumerator() | ForEach-Object {
    $x = $_.Key
    Write-Output "<<<...................................................."
    Write-Output "Updating Field:"
    Write-Output  $_.Key
    Write-Output "TO:"
    write-output $hashtable.$x  
    Write-Output " "
    try {
        Set-CrmRecord -EntityLogicalName organization -Id $orgid -Fields @{$_.Key = $_.Value} -ErrorAction Stop
        Write-Output "$($_.Key.ToUpper()) UPDATE SUCCESS!"
        Write-Output "....................................................>>>"
        Write-Output " "
        Write-Output " "
    }
    catch {
        throw $Error[0].Exception
    }
}

