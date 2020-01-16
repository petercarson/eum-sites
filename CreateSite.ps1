Param
(
    [Parameter (Mandatory = $false)][int]$listItemID
)

$Global:AzureAutomation = (Get-Command "Get-AutomationVariable" -ErrorAction SilentlyContinue)
if ($AzureAutomation) { 
    . .\EUMSites_Helper.ps1
}
else {
    $DistributionFolder = (Split-Path $MyInvocation.MyCommand.Path)
    . $DistributionFolder\EUMSites_Helper.ps1
}

LoadEnvironmentSettings

Connect-PnPOnline -Url $SitesListSiteURL -Credentials $SPCredentials -CreateDrive

New-Item -Path $pnpTemplatePath -ItemType "directory" -Force | out-null
Copy-Item -Path "spo:.\pnptemplates\*" -Destination $pnpTemplatePath -Force
Write-Verbose -Verbose -Message "Templates:"
Get-ChildItem $pnpTemplatePath | ForEach-Object { Write-Verbose -Verbose -Message $_.Name }

if ($listItemID -gt 0) {
    # Get the specific Site Collection List item in master site for the site that needs to be created

    $pendingSiteCollections = Get-PnPListItem -List $SiteListName -Query "
    <View>
        <Query>
            <Where>
                <And>
                    <Eq>
                        <FieldRef Name='ID'/>
                        <Value Type='Integer'>$listItemID</Value>
                    </Eq>
                    <IsNull>
                        <FieldRef Name='EUMSiteCreated'/>
                    </IsNull>
                </And>
            </Where>
        </Query>
        <ViewFields>
            <FieldRef Name='ID'></FieldRef>
            <FieldRef Name='Title'></FieldRef>
            <FieldRef Name='EUMSiteURL'></FieldRef>
            <FieldRef Name='EUMAlias'></FieldRef>
            <FieldRef Name='EUMPublicGroup'></FieldRef>
            <FieldRef Name='EUMBreadcrumbHTML'></FieldRef>
            <FieldRef Name='EUMParentURL'></FieldRef>
            <FieldRef Name='EUMSiteTemplate'></FieldRef>
            <FieldRef Name='EUMDivision'></FieldRef>
            <FieldRef Name='EUMCreateTeam'></FieldRef>
            <FieldRef Name='Author'></FieldRef>
        </ViewFields>
    </View>"
}
else {
    # Check the Site Collection List in master site for any sites that need to be created

    $pendingSiteCollections = Get-PnPListItem -List $SiteListName -Query "
    <View>
        <Query>
            <Where>
                <IsNull>
                    <FieldRef Name='EUMSiteCreated'/>
                </IsNull>
            </Where>
        </Query>
        <ViewFields>
            <FieldRef Name='ID'></FieldRef>
            <FieldRef Name='Title'></FieldRef>
            <FieldRef Name='EUMSiteURL'></FieldRef>
            <FieldRef Name='EUMAlias'></FieldRef>
            <FieldRef Name='EUMPublicGroup'></FieldRef>
            <FieldRef Name='EUMBreadcrumbHTML'></FieldRef>
            <FieldRef Name='EUMParentURL'></FieldRef>
            <FieldRef Name='EUMSiteTemplate'></FieldRef>
            <FieldRef Name='EUMDivision'></FieldRef>
            <FieldRef Name='EUMCreateTeam'></FieldRef>
            <FieldRef Name='Author'></FieldRef>
        </ViewFields>
    </View>"
}


if ($pendingSiteCollections.Count -gt 0) {
    # Get the time zone of the master site
    $spWeb = Get-PnPWeb -Includes RegionalSettings.TimeZone
    [int]$timeZoneId = $spWeb.RegionalSettings.TimeZone.Id

    # Iterate through the pending sites. Create them if needed, and apply template
    $pendingSiteCollections | ForEach {
        $pendingSite = $_

        [string]$siteTitle = $pendingSite["Title"]
        [string]$alias = $pendingSite["EUMAlias"]
        if ($alias) {
            # Replace spaces in Alias with dashes
            $alias = $alias -replace '\s', '-'
            $siteURL = "$($WebAppURL)/sites/$alias"
        }
        else {
            [string]$siteURL = "$WebAppURL/sites/$($pendingSite['EUMSiteURL'])"
        }
        [boolean]$publicGroup = $pendingSite["EUMPublicGroup"]
        [string]$breadcrumbHTML = $pendingSite["EUMBreadcrumbHTML"]
        [string]$parentURL = $pendingSite["EUMParentURL"]
        [string]$Division = $pendingSite["EUMDivision"].LookupValue
        [string]$eumSiteTemplate = $pendingSite["EUMSiteTemplate"].LookupValue
        [boolean]$eumCreateTeam = $pendingSite["EUMCreateTeam"]
        [string]$author = $pendingSite["Author"].Email

        if ($parentURL -eq "") {
            $divisionSiteURL = Get-PnPListItem -List "Divisions" -Query "
																<View>
																	<Query>
																		<Where>
																			<Eq>
																				<FieldRef Name='Title'/>
																				<Value Type='Text'>$Division</Value>
																			</Eq>
																		</Where>
																	</Query>
																	<ViewFields>
																		<FieldRef Name='Title'></FieldRef>
																		<FieldRef Name='SiteURL'></FieldRef>
																	</ViewFields>
																</View>"
		
            if ($divisionSiteURL.Count -eq 1) {
                $parentURL = $divisionSiteURL["SiteURL"].Url
            }
        }

        $siteTemplate = Get-PnPListItem -List "SiteTemplates" -Query "
													<View>
														<Query>
															<Where>
																<Eq>
																	<FieldRef Name='Title'/>
																	<Value Type='Text'>$eumSiteTemplate</Value>
																</Eq>
															</Where>
														</Query>
														<ViewFields>
															<FieldRef Name='Title'></FieldRef>
															<FieldRef Name='BaseClassicSiteTemplate'></FieldRef>
															<FieldRef Name='BaseModernSiteType'></FieldRef>
															<FieldRef Name='PnPSiteTemplate'></FieldRef>
														</ViewFields>
													</View>"
		
        $baseSiteTemplate = ""
        $baseSiteType = ""
        $pnpSiteTemplate = ""
        $siteCreated = $false

        if ($siteTemplate.Count -eq 1) {
            $baseSiteTemplate = $siteTemplate["BaseClassicSiteTemplate"]
            $baseSiteType = $siteTemplate["BaseModernSiteType"]
            if ($siteTemplate["PnPSiteTemplate"] -ne $null) {
                $pnpSiteTemplate = "$pnpTemplatePath\$($siteTemplate["PnPSiteTemplate"])"
            }
        }

        # Classic style sites
        if ($baseSiteTemplate) {
            # Create the site
            if ($siteCollection) {
                # Create site (if it exists, it will error but not modify the existing site)
                Write-Verbose -Verbose -Message "Creating site collection $($siteURL) with base template $($baseSiteTemplate). Please wait..."
                try {
                    New-PnPTenantSite -Title $siteTitle -Url $siteURL -Owner $author -TimeZone $timeZoneId -Template $baseSiteTemplate -RemoveDeletedSite -Wait -Force -ErrorAction Stop
                }
                catch { 
                    Write-Error "Failed creating site collection $($siteURL)"
                    Write-Error $_
                }
            }
            else {
                # Connect to parent site
                Helper-Connect-PnPOnline -Url $parentURL

                # Create the subsite
                Write-Verbose -Verbose -Message "Creating subsite $($siteURL) with base template $($baseSiteTemplate) under $($parentURL). Please wait..."

                [string]$subsiteURL = $siteURL.Replace($parentURL, "").Trim('/')
                New-PnPWeb -Title $siteTitle -Url $subsiteURL -Template $baseSiteTemplate

                Disconnect-PnPOnline
            }
            $siteCreated = $true

        }
        # Modern style sites
        else {
            # Create the site
            switch ($baseSiteType) {
                "CommunicationSite" {
                    try {
                        Write-Verbose -Verbose -Message "Creating site collection $($siteURL) with modern type $($baseSiteType). Please wait..."
                        $siteURL = New-PnPSite -Type CommunicationSite -Title $siteTitle -Url $siteURL -ErrorAction Stop
                        $siteCreated = $true
                    }
                    catch { 
                        Write-Error "Failed creating site collection $($siteURL)"
                        Write-Error $_
                        exit
                    }
                }
                "TeamSite" {
                    try {
                        Write-Verbose -Verbose -Message "Creating site collection $($siteURL) with modern type $($baseSiteType). Please wait..."
                        if ($publicGroup) {
                            $siteURL = New-PnPSite -Type TeamSite -Title $siteTitle -Alias $alias -IsPublic -ErrorAction Stop
                        }
                        else {
                            $siteURL = New-PnPSite -Type TeamSite -Title $siteTitle -Alias $alias -ErrorAction Stop
                        }
                        $siteCreated = $true

                        
                    }
                    catch { 
                        Write-Error "Failed creating site collection $($siteURL)"
                        Write-Error $_
                        exit
                    }
                }
            }
        }

        if ($siteCreated) {
            Helper-Connect-PnPOnline -Url $AdminURL
            $spSite = Get-PnPTenantSite -Url $siteURL
            $retries = 0

            while (($spSite.Status -ne "Active") -and ($retries -lt 120)) {
                Start-Sleep -Seconds 60
                $retries += 1
                $spSite = Get-PnPTenantSite -Url $siteURL
            }
            Disconnect-PnPOnline

            $groupId = $spSite.GroupId

            # add the requester to the Owners group
            if ($groupId) {
                AddGroupOwner -groupID $groupId -email $author
            }            

            Helper-Connect-PnPOnline -Url $siteURL
            # Set the site collection admin
            if ($SiteCollectionAdministrator -ne "") {
                Add-PnPSiteCollectionAdmin -Owners $SiteCollectionAdministrator
            }
            Add-PnPSiteCollectionAdmin -Owners $author

            if ($pnpSiteTemplate) {
                $retries = 0
                $pnpTemplateApplied = $false
                while (($retries -lt 20) -and ($pnpTemplateApplied -eq $false)) {
                    Start-Sleep -Seconds 30
                    Write-Verbose -Verbose -Message "Applying template $($pnpSiteTemplate) Please wait..."
                    try {
                        $retries += 1
                        Set-PnPTraceLog -On -Level Debug
                        Apply-PnPProvisioningTemplate -Path $pnpSiteTemplate -ErrorAction Stop
                        $pnpTemplateApplied = $true
                    }
                    catch {      
                        Write-Verbose -Verbose -Message "Failed applying PnP template."
                        Write-Verbose -Verbose -Message $_
                    }
                }
            }
            
            # Create the team if needed
            if ($eumCreateTeam) {
                Write-Verbose -Verbose -Message "Creating Microsoft Team"

                Connect-MicrosoftTeams -Credential $SPCredentials
                $team = New-Team -GroupId $groupId
                $teamsChannels = Get-TeamChannel -GroupId $groupId
                $generalChannel = $teamsChannels | Where-Object { $_.DisplayName -eq 'General' }
                $generalChannelId = $generalChannel.Id
                Disconnect-MicrosoftTeams

                Write-Verbose -Verbose -Message "groupId = $($groupId), generalChannelId = $($generalChannelId)"
                AddOneNoteTeamsChannelTab -groupId $groupId -channelName 'General' -teamsChannelId $generalChannelId -siteURL $siteURL
                AddTeamsChannelRequestFormToChannel -groupId $groupId -teamsChannelId $generalChannelId
            }

            # Reconnect to the master site and update the site collection list
            Helper-Connect-PnPOnline -Url $SitesListSiteURL

            # Set the breadcrumb HTML
            [string]$breadcrumbHTML = GetBreadcrumbHTML -siteURL $siteURL -siteTitle $siteTitle -parentURL $parentURL

            # Set the site created date, breadcrumb, and site URL
            [Microsoft.SharePoint.Client.ListItem]$spListItem = Set-PnPListItem -List $SiteListName -Identity $pendingSite.Id -Values @{ "EUMSiteCreated" = [System.DateTime]::Now; "EUMBreadcrumbHTML" = $breadcrumbHTML; "EUMSiteURL" = $siteURL; "EUMParentURL" = $parentURL }
        }

        # Reconnect to the master site for the next iteration
        Helper-Connect-PnPOnline -Url $SitesListSiteURL
    }
}
else {
    Write-Verbose -Verbose -Message "No sites pending creation"
}