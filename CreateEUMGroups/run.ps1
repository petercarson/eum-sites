if ($Env:POSTMethod)
{
    # POST method: $req
    $requestBody = Get-Content $req -Raw | ConvertFrom-Json
    $ID = $requestBody.id
}

[string]$DistributionFolder = $Env:distributionFolder

if (-not $DistributionFolder)
{
    $DistributionFolder = (Split-Path $MyInvocation.MyCommand.Path)
    $DistributionFolderArray = $DistributionFolder.Split('\')
    $DistributionFolderArray[$DistributionFolderArray.Count - 1] = ""
    $DistributionFolder = $DistributionFolderArray -join "\"
}

. $DistributionFolder\EUMSites_Helper.ps1
LoadEnvironmentSettings

if ($listItemID)
{
    # Get the specific Site Collection List item in master site for the site that needs to be created
    Helper-Connect-PnPOnline -Url $SitesListSiteURL

    $pendingSiteCollections = Get-PnPListItem -List $SiteListName -Query "
    <View>
        <Query>
            <Where>
                <And>
                    <Eq>
                        <FieldRef Name='ID'/>
                        <Value Type='Integer'>$itemId</Value>
                    </Eq>
                    <IsNotNull>
                        <FieldRef Name='EUMEUMGroup'/>
                    </IsNotNull>
                    <And>
                        <IsNotNull>
                            <FieldRef Name='EUMEUMPermission'/>
                        </IsNotNull>
                        <IsNull>
                            <FieldRef Name='EUMEUMGroupCreated'/>
                        </IsNull>
                    </And>
                </And>
            </Where>
        </Query>
        <ViewFields>
            <FieldRef Name='ID'></FieldRef>
            <FieldRef Name='Title'></FieldRef>
            <FieldRef Name='EUMSiteURL'></FieldRef>
            <FieldRef Name='EUMEUMGroup'></FieldRef>
            <FieldRef Name='EUMEUMPermission'></FieldRef>
            <FieldRef Name='EUMSiteTemplate'></FieldRef>
        </ViewFields>
    </View>"
}
else
{
    # Check the Site Collection List in master site for any sites that need to be created
    Helper-Connect-PnPOnline -Url $SitesListSiteURL

    $pendingSiteCollections = Get-PnPListItem -List $SiteListName -Query "
    <View>
        <Query>
            <Where>
                <And>
                    <IsNotNull>
                        <FieldRef Name='EUMEUMGroup'/>
                    </IsNotNull>
                    <And>
                        <IsNotNull>
                            <FieldRef Name='EUMEUMPermission'/>
                        </IsNotNull>
                        <IsNull>
                            <FieldRef Name='EUMEUMGroupCreated'/>
                        </IsNull>
                    </And>
                </And>
            </Where>
        </Query>
        <ViewFields>
            <FieldRef Name='ID'></FieldRef>
            <FieldRef Name='Title'></FieldRef>
            <FieldRef Name='EUMSiteURL'></FieldRef>
            <FieldRef Name='EUMEUMGroup'></FieldRef>
            <FieldRef Name='EUMEUMPermission'></FieldRef>
            <FieldRef Name='EUMSiteTemplate'></FieldRef>
        </ViewFields>
    </View>"
}

if ($pendingSiteCollections.Count -gt 0)
{
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

    # Iterate through the pending sites. Create the groups if needed, and update permissions
    $pendingSiteCollections | ForEach {
        $groupCreated = $false
        $pendingSite = $_

        $content = @{
            "Domain_FK"= $Domain_FK;
            "RoleStatus_FK"= 1;
            "SystemConfiguration_FK"= $SystemConfiguration_FK;
            "RoleName"= $pendingSite["EUMEUMGroup"];
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
            [Microsoft.SharePoint.Client.ListItem]$spListItem = Set-PnPListItem -List $SiteListName -Identity $pendingSite.Id -Values @{ "EUMEUMGroupCreated" = [System.DateTime]::Now }

            $groupName = $pendingSite["EUMEUMGroup"]

            # Enable external sharing
            Connect-PnPOnline -url ($pendingSite["EUMSiteURL"]).Url -Credentials $credentials
            # Possible values Disabled, ExternalUserSharingOnly, ExternalUserAndGuestSharing, ExistingExternalUserSharingOnly
            Set-PnPTenantSite -Url ($pendingSite["EUMSiteURL"]).Url -Sharing ExternalUserAndGuestSharing

            # Add the appropriate permission to the site for the group
            Set-PnPWebPermission -User $groupName -AddRole $pendingSite["EUMEUMPermission"]

            switch ($eumSiteTemplate)
            {
                "Modern Client Site"
                    {
                    # Remove the rights to the Private Documents library
                    $list = Get-PnPList -Identity "Private Documents"
                    $list.BreakRoleInheritance($true, $true)

                    Set-PnPListPermission -Identity "Private Documents" -User $groupName -RemoveRole $pendingSite["EUMEUMPermission"]
                    }
            }

            # Update the permissions to the Site Collection List in master site to give the group read access
            Connect-PnPOnline -Url $SitesListSiteURL -Credentials $credentials

            Set-PnPListItemPermission -List $SiteListName -Identity $pendingSite["ID"] -User $groupName -AddRole "Read"
        }
    }
}
