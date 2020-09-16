Param
(
    [Parameter (Mandatory = $true)][int]$listItemID
)

$Global:AzureAutomation = (Get-Command "Get-AutomationVariable" -ErrorAction SilentlyContinue)
if ($AzureAutomation) { 
    . .\..\EUMSites_Helper.ps1
}
else {
    . $PSScriptRoot\EUMSites_Helper.ps1
}

LoadEnvironmentSettings

# Get the specific Site Collection List item in master site for the site that needs to be created
$connLanding = Helper-Connect-PnPOnline -Url $SitesListSiteURL

$pendingSiteCollections = Get-PnPListItem -Connection $connLanding -List $SiteListName -Query "
<View>
    <Query>
        <Where>
            <Eq>
                <FieldRef Name='ID'/>
                <Value Type='Integer'>$listItemID</Value>
            </Eq>
        </Where>
    </Query>
    <ViewFields>
        <FieldRef Name='ID'></FieldRef>
        <FieldRef Name='Title'></FieldRef>
        <FieldRef Name='EUMSiteURL'></FieldRef>
    </ViewFields>
</View>"

if ($pendingSiteCollections.Count -gt 0) {
    Add-Type -path "$DistributionFolder\DLLs\IdentityModel.dll"
    Add-Type -path "$DistributionFolder\DLLs\Newtonsoft.Json.dll"
    Add-Type -path "$DistributionFolder\DLLs\System.ValueTuple.dll"

    $disco = [IdentityModel.Client.DiscoveryClient]::GetAsync($EUMURL + "/idsrv").GetAwaiter().GetResult()
    if ($disco.IsError) {
        Write-Output $disco.Error
    }

    $tokenClient = New-Object IdentityModel.Client.TokenClient($disco.TokenEndpoint, $Global:EUMClientID, $Global:EUMSecret)
    $cancelToken = New-Object System.Threading.CancellationToken
    $tokenResponse = [IdentityModel.Client.TokenClientExtensions]::RequestClientCredentialsAsync($tokenClient, "extranet_api_v4", $null, $cancelToken).GetAwaiter().GetResult()

    if ($tokenResponse.IsError) {
        Write-Output $tokenResponse.Error
    }

    $client = New-Object System.Net.Http.HttpClient
    [System.Net.Http.HttpClientExtensions]::SetBearerToken($client, $tokenResponse.AccessToken)

    # Iterate through the pending sites. Create the groups if needed, and update permissions
    $pendingSiteCollections | ForEach {
        $groupCreated = $false
        $pendingSite = $_
        $groupName = $groupName = "EUM " + $pendingSite["Title"]
        $siteURL = $pendingSite["EUMSiteURL"]

        $content = @{
            "Domain_FK"                = $Domain_FK
            "RoleStatus_FK"            = 1
            "SystemConfiguration_FK"   = $SystemConfiguration_FK
            "RoleName"                 = $groupName
            "AvailableForRegistration" = $true
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

        if ($groupCreated) {
            # Enable external sharing
            $connSite = Helper-Connect-PnPOnline -url $siteURL
            # Possible values Disabled, ExternalUserSharingOnly, ExternalUserAndGuestSharing, ExistingExternalUserSharingOnly
            Set-PnPTenantSite -Url $siteURL -Sharing ExternalUserAndGuestSharing -Connection $connSite

            # Break permissions inheritance on the Private Documents library
            $list = Get-PnPList -Identity "Private Documents" -Connection $connSite
            $list.BreakRoleInheritance($true, $true)

            # Add the appropriate permission to the site for the group
            Set-PnPWebPermission -User $groupName -AddRole "Read" -Connection $connSite
        }
    }
}
