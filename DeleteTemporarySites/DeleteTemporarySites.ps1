Param
(
    [Parameter (Mandatory = $false)][int]$listItemID
)

$Global:AzureAutomation = (Get-Command "Get-AutomationVariable" -ErrorAction SilentlyContinue)
if ($AzureAutomation) { 
    $DistributionFolder = Get-Location

    # Get automation variables and credentials
    $Global:storageName = Get-AutomationVariable -Name 'AzureStorageName'
    $Global:credentialName = Get-AutomationVariable -Name 'AutomationCredentialName'
    $Global:connectionString = Get-AutomationPSCredential -Name 'AzureStorageConnectionString'
    $Global:SPCredentials = Get-AutomationPSCredential -Name $credentialName

    # Get EUMSites_Helper.ps1 and sharepoint.config from Azure storage
    $Global:storageContext = New-AzureStorageContext -ConnectionString $connectionString.GetNetworkCredential().Password
    Get-AzureStorageFileContent -ShareName $storageName -Path "sharepoint.config" -Context $storageContext -Force
    Get-AzureStorageFileContent -ShareName $storageName -Path "EUMSites_Helper.ps1" -Context $storageContext -Force

    # Get site templates and branding files from azure storage
    New-Item -ItemType Directory -Path "$($DistributionFolder)\SiteTemplates"
    Get-AzureStorageFile -ShareName $storageName -Path "SiteTemplates" -Context $storageContext | Get-AzureStorageFile | ? {$_.GetType().Name -eq "CloudFile"} | Get-AzureStorageFileContent -Force -Destination "$($DistributionFolder)\SiteTemplates"

		New-Item -ItemType Directory -Path "$($DistributionFolder)\SiteTemplates\Pages"
		Get-AzureStorageFile -ShareName $storageName -Path "SiteTemplates\Pages" -Context $storageContext | Get-AzureStorageFile | ? {$_.GetType().Name -eq "CloudFile"} | Get-AzureStorageFileContent -Force -Destination "$($DistributionFolder)\SiteTemplates\Pages"
}
else {
    $DistributionFolder = (Split-Path $MyInvocation.MyCommand.Path)
    $DistributionFolderArray = $DistributionFolder.Split('\')
    $DistributionFolderArray[$DistributionFolderArray.Count - 1] = ""
    $DistributionFolder = $DistributionFolderArray -join "\"
}

. $DistributionFolder\EUMSites_Helper.ps1
LoadEnvironmentSettings

# Check the Site Collection List in master site for any sites that need to be created
Helper-Connect-PnPOnline -Url $SitesListSiteURL

$temporarySiteCollections = Get-PnPListItem -List $SiteListName -Query "
<View>
    <Query>
        <Where>
            <BeginsWith>
				<FieldRef Name='EUMParentURL'/>
				<Value Type='Url'>/sites/tempdemos</Value>
            </BeginsWith>
        </Where>
    </Query>
    <ViewFields>
        <FieldRef Name='ID'></FieldRef>
        <FieldRef Name='Title'></FieldRef>
        <FieldRef Name='EUMSiteURL'></FieldRef>
        <FieldRef Name='EUMAlias'></FieldRef>
    </ViewFields>
</View>"


if ($temporarySiteCollections.Count -gt 0) {
    # Iterate through the pending sites. Create them if needed, and apply template
    $temporarySiteCollections | ForEach {
        $temporarySite = $_

        if ($temporarySite["EUMAlias"] -ne "")
        {
            
        }
    }
}
else {
    Write-Output "No temporary sites to delete"
}