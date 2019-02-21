#Param
#(
#    [Parameter (Mandatory = $true)][string]$siteURL,
#    [Parameter (Mandatory = $true)][string]$siteTitle
#)

$siteURL = "https://eumdemo.sharepoint.com/sites/webinardemo4"
$siteTitle = "Webinar Demo 4"

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
}
else {
    $DistributionFolder = (Split-Path $MyInvocation.MyCommand.Path)
    $DistributionFolderArray = $DistributionFolder.Split('\')
    $DistributionFolderArray[$DistributionFolderArray.Count - 1] = ""
    $DistributionFolder = $DistributionFolderArray -join "\"
}

. $DistributionFolder\EUMSites_Helper.ps1
LoadEnvironmentSettings

# Get the specific Site Collection List item in master site for the site with the EUM group to be created
Helper-Connect-PnPOnline -Url $SitesListSiteURL

$pendingSiteCollections = Get-PnPListItem -List $SiteListName -Query "
<View>
    <Query>
        <Where>
            <Eq>
                <FieldRef Name='EUMSiteURL'/>
                <Value Type='String'>$siteURL</Value>
            </Eq>
        </Where>
    </Query>
    <ViewFields>
        <FieldRef Name='EUMSiteURL'></FieldRef>
    </ViewFields>
</View>"

if ($pendingSiteCollections.Count -eq 0)
{
    $EUMGroup = "EUM - " + $siteTitle
    $EUMPermission = "Read"

    Add-Type -path "$DistributionFolder\CreateEUMGroups\IdentityModel.dll"
    Add-Type -path "$DistributionFolder\CreateEUMGroups\Newtonsoft.Json.dll"
    Add-Type -path "$DistributionFolder\CreateEUMGroups\System.ValueTuple.dll"

    $disco = [IdentityModel.Client.DiscoveryClient]::GetAsync($EUMURL + "/idsrv").GetAwaiter().GetResult()
    if ($disco.IsError) {
        Write-Output $disco.Error
    }

    $tokenClient = New-Object IdentityModel.Client.TokenClient($disco.TokenEndpoint, $EUMClientID, $EUMSecret)
    $cancelToken = New-Object System.Threading.CancellationToken
    $tokenResponse = [IdentityModel.Client.TokenClientExtensions]::RequestClientCredentialsAsync($tokenClient, "extranet_api_v4", $null, $cancelToken).GetAwaiter().GetResult()

    if ($tokenResponse.IsError) {
        Write-Output $tokenResponse.Error
    }

    $client = New-Object System.Net.Http.HttpClient
    [System.Net.Http.HttpClientExtensions]::SetBearerToken($client, $tokenResponse.AccessToken)

    $groupCreated = $false

    $content = @{
        "Domain_FK"= $Domain_FK;
        "RoleStatus_FK"= 1;
        "SystemConfiguration_FK"= $SystemConfiguration_FK;
        "RoleName"= $EUMGroup;
        "AvailableForRegistration"= $true
    }

    $json = ConvertTo-Json $content
    $stringContent = New-Object System.Net.Http.StringContent($json, [System.Text.Encoding]::UTF8, "application/json");
    $response = $client.PostAsync($EUMURL + "/_API/v4/Roles", $stringContent).GetAwaiter().GetResult()

    if ($response.IsSuccessStatusCode) {
        $content = $response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
        $content
        $conv = ConvertFrom-Json($content)
        $conv
        $groupCreated = $true
    }
    else {
        Write-Output $response.StatusCode
    }

    if ($groupCreated)
    {
        # Set the Group Created value
        [Microsoft.SharePoint.Client.ListItem]$spListItem = Add-PnPListItem -List $SiteListName -Values @{ "EUMEUMGroup" = $EUMGroup; "EUMEUMGroupCreated" = [System.DateTime]::Now; "EUMEUMPermission" = "Read"; "EUMSiteURL" = $siteURL; "Title" = $siteTitle }

        # Enable external sharing
        Connect-PnPOnline -url $siteURL -Credentials $credentials
        # Possible values Disabled, ExternalUserSharingOnly, ExternalUserAndGuestSharing, ExistingExternalUserSharingOnly
        Set-PnPTenantSite -Url $siteURL -Sharing ExternalUserAndGuestSharing

        Start-Sleep -Seconds 90

        Set-PnPTraceLog -On -Level Debug
        $pnpSiteTemplate = $DistributionFolder + "\SiteTemplates\Project-Template-Template.xml"
        Apply-PnPProvisioningTemplate -Path $pnpSiteTemplate

        # Remove the rights to the Shared Documents library
        $list = Get-PnPList -Identity "Shared Documents"
        $list.BreakRoleInheritance($true, $true)

        # Add the appropriate permission to the site for the group
        Set-PnPWebPermission -User $EUMGroup -AddRole $EUMPermission

        # Update the permissions to the Site Collection List in master site to give the group read access
        Connect-PnPOnline -Url $SitesListSiteURL -Credentials $credentials

        Set-PnPListItemPermission -List $SiteListName -Identity $spListItem.Id -User $EUMGroup -AddRole "Read"
    }
}
