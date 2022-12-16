param
( 
    [Parameter(Mandatory = $true, HelpMessage = "clientId")] [String]$clientId,
    [Parameter(Mandatory = $true, HelpMessage = "clientSecret")] [String]$clientSecret,
    [Parameter(Mandatory = $true, HelpMessage = "crmURL")] [String]$crmURL,
    [Parameter(Mandatory = $true, HelpMessage = "Tenant ID")] [String]$tenantId
)
# import the file of useful re-usable functions
. $PSScriptRoot\Modules\Functions.ps1
# install and import modules
InstallAndImport -PSModules "Microsoft.Xrm.Tooling.CrmConnector.PowerShell"
# get access token
$newToken = XRM-APIgetAccessToken -clientId $clientId -clientSecret $clientSecret -crmURL $crmURL -tenantId $tenantId
# grab the system settings from the source *donor" environment - uses Invoke-RestMethod to get a PSObject from the json
$orgObject = (XRM-APIqueryEntity -crmURL $crmURL -entityName "organization" -authResponse $newToken).value
$orgObject |  ForEach-Object { 
    # remove some stuff we don't need from the response object
    # some of these are read only..
    # some are irrelevant to our needs
    $_.PSObject.Properties.Remove("@odata.etag") # dont want to migrate
    $_.PSObject.Properties.Remove("createdon") # dont want to migrate
    $_.PSObject.Properties.Remove("organizationid") # dont want to migrate
    $_.PSObject.Properties.Remove("_basecurrencyid_value") # not sure if i can mess with this one
    $_.PSObject.Properties.Remove("name") # dont want to migrate
    $_.PSObject.Properties.Remove("businessclosurecalendarid") # dont want to migrate id can change
    $_.PSObject.Properties.Remove("_defaultemailserverprofileid_value") # dont want to migrate
    $_.PSObject.Properties.Remove("integrationuserid")   # dont want to migrate
    $_.PSObject.Properties.Remove("systemuserid") # dont want to migrate
    $_.PSObject.Properties.Remove("delegatedadminuserid") # dont want to migrate
    $_.PSObject.Properties.Remove("supportuserid") # dont want to migrate
    $_.PSObject.Properties.Remove("versionnumber") # dont want to migrate
    $_.PSObject.Properties.Remove("modifiedon") # dont want to migrate
    $_.PSObject.Properties.Remove("_modifiedby_value") # dont want to migrate
    $_.PSObject.Properties.Remove("_createdby_value") # dont want to migrate
    $_.PSObject.Properties.Remove("_modifiedonbehalfby_value") # dont want to migrate
    $_.PSObject.Properties.Remove("_createdonbehalfby_value") # dont want to migrate
    $_.PSObject.Properties.Remove("_acknowledgementtemplateid_value") # dont want to migrate
    $_.PSObject.Properties.Remove("releasewavename") # dont want to migrate
    $_.PSObject.Properties.Remove("requireapprovalforuseremail") # <-- throws an error with this enabled. but its my permissions in the tenant.
    $_.PSObject.Properties.Remove("requireapprovalforqueueemail") # <-- throws an error with this enabled. but its my permissions in the tenant.
    # and we will remove anything with a NULL/Empty value and those that are an empty string.
    # these either haven't been set or default to null and we dont want NULL values anyway 
    $NonEmptyProperties = $_.psobject.Properties | Where-Object { $null -ne $_.Value -and $_.Value -ne ""  } | Select-Object -ExpandProperty Name
    # Then convert this back to json that we can use to SET these values in target environments
    $newObj = $_ | Select-Object -Property $NonEmptyProperties | ConvertTo-Json
} 
# create json filename from crmURL parameter # and a foldername to sling it in
$ExportFileName = $crmURL.Replace("https://", "") | % { $_.Substring(0, $_.IndexOf('.')) + ".json" } 
$jsonfolder = "jsonExport"
# check the path exists and drop the file in there.. if it dont exist create it
if (Test-Path -Path $jsonfolder) {
    Write-Output "Exporting Org Settings"
    $newObj | Out-File .\$jsonfolder\$ExportFileName
    # get the path and upload as an artifact
    $artifact = (Get-ChildItem -Path .\$jsonFolder\).FullName.ToString()
    Write-Host "##vso[artifact.upload containerfolder=$jsonFolder;artifactname=orgSettings]$artifact"  
} else {
    Write-Output "Creating directory"
    mkdir $jsonfolder
    Write-Output "Exporting Org Settings"
    $orgObject | Out-File .\$jsonfolder\$ExportFileName
    # get the path and upload as an artifact
    $artifact = (Get-ChildItem -Path .\$jsonFolder\).FullName.ToString()
    Write-Host "##vso[artifact.upload containerfolder=$jsonFolder;artifactname=orgSettings]$artifact"      
}