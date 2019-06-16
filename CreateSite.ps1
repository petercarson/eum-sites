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

Helper-Connect-PnPOnline -Url $SitesListSiteURL

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
            $siteURL = "$($WebAppURL)/sites/$alias"
        }
        else {
            [string]$siteURL = "$WebAppURL/sites/$($pendingSite['EUMSiteURL'])"
        }
        [string]$publicGroup = $pendingSite["EUMPublicGroup"]
        [string]$breadcrumbHTML = $pendingSite["EUMBreadcrumbHTML"]
        [string]$parentURL = $pendingSite["EUMParentURL"]
        [string]$Division = $pendingSite["EUMDivision"].LookupValue
        [string]$eumSiteTemplate = $pendingSite["EUMSiteTemplate"]
        [string]$eumCreateTeam = $pendingSite["EUMCreateTeam"]
        [string]$author = $pendingSite["Author"].Email

		if ($parentURL -eq "")
		{
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
		
            if ($divisionSiteURL.Count -eq 1)
            {
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

        if ($siteTemplate.Count -eq 1)
        {
			$baseSiteTemplate = $siteTemplate["BaseClassicSiteTemplate"]
			$baseSiteType = $siteTemplate["BaseModernSiteType"]
            if ($siteTemplate["PnPSiteTemplate"] -ne $null)
            {
    			$pnpSiteTemplate = "$pnpTemplatePath\$($siteTemplate["PnPSiteTemplate"])"
            }
        }

        # Classic style sites
        if ($baseSiteTemplate) {
            # Create the site
            if ($siteCollection) {
                # Create site (if it exists, it will error but not modify the existing site)
                Write-Output "Creating site collection $($siteURL) with base template $($baseSiteTemplate). Please wait..."
                try {
                    New-PnPTenantSite -Title $siteTitle -Url $siteURL -Owner $author -TimeZone $timeZoneId -Template $baseSiteTemplate -RemoveDeletedSite -Wait -Force -ErrorAction Stop
                }
                catch { 
                    Write-Error "Failed creating site collection $($siteURL)"
                    Write-Error $_
                    exit
                }
            }
            else {
                # Connect to parent site
                Helper-Connect-PnPOnline -Url $parentURL

                # Create the subsite
                Write-Output "Creating subsite $($siteURL) with base template $($baseSiteTemplate) under $($parentURL). Please wait..."

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
                        Write-Output "Creating site collection $($siteURL) with modern type $($baseSiteType). Please wait..."
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
                        Write-Output "Creating site collection $($siteURL) with modern type $($baseSiteType). Please wait..."
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
            Helper-Connect-PnPOnline -Url $siteURL

            # Set the site collection admin
            if ($SiteCollectionAdministrator -ne "")
            {
                Add-PnPSiteCollectionAdmin -Owners $SiteCollectionAdministrator
            }

            if ($pnpSiteTemplate) {
                # Pause the script to allow time for the modern site to finish provisioning
                Write-Output "Pausing for 300 seconds. Please wait..."
                Start-Sleep -Seconds 300

                Write-Output "Applying template $($pnpSiteTemplate) Please wait..."

                try {
		            Set-PnPTraceLog -On -Level Debug
                    Apply-PnPProvisioningTemplate -Path $pnpSiteTemplate -ErrorAction Stop
                }
                catch {
                    Write-Error "Failed applying PnP template."
                    Write-Error $_
                    exit
                }
            }
            
            If (($pnpSiteTemplate -like "*Client-Template.xml"))
            {
                Add-PnPFolder -Name "Quotes" -Folder "/Shared Documents"
                Add-PnPFolder -Name "Signed Quotes" -Folder "/Shared Documents/Quotes"
                Add-PnPFolder -Name "Invoices" -Folder "/Shared Documents"

                Add-PnPFolder -Name "Business Development" -Folder "/Private Documents"
                Add-PnPFolder -Name "Confidential" -Folder "/Private Documents"
                Add-PnPFolder -Name "Quotes" -Folder "/Private Documents"

                Remove-PnPContentTypeFromList -List "Shared Documents" -ContentType "Document"
                Remove-PnPContentTypeFromList -List "Private Documents" -ContentType "Document"
            }

            # Create the team if needed
            if ($eumCreateTeam -eq $true)
            {
                Write-Output "Creating Microsoft Team"
                Helper-Connect-PnPOnline -Url $AdminURL
                $spSite = Get-PnPTenantSite -Url $siteURL
                $groupId = $spSite.GroupId

                Connect-MicrosoftTeams -Credential $SPCredentials
                $team = New-Team -GroupId $groupId
                Disconnect-MicrosoftTeams
                Disconnect-PnPOnline
            }

            # Reconnect to the master site and update the site collection list
            Helper-Connect-PnPOnline -Url $SitesListSiteURL

            # Set the breadcrumb HTML
            [string]$breadcrumbHTML = GetBreadcrumbHTML -siteURL $siteURL -siteTitle $siteTitle -parentURL $parentURL

            # Set the site created date, breadcrumb, and site URL
            [Microsoft.SharePoint.Client.ListItem]$spListItem = Set-PnPListItem -List $SiteListName -Identity $pendingSite.Id -Values @{ "EUMSiteCreated" = [System.DateTime]::Now; "EUMBreadcrumbHTML" = $breadcrumbHTML; "EUMSiteURL" = $siteURL; "EUMParentURL" = $parentURL }

            # Install Masthead on the site
            # Install-To-Site $siteURL
        }

        # Reconnect to the master site for the next iteration
        Helper-Connect-PnPOnline -Url $SitesListSiteURL
    }
}
else {
    Write-Output "No sites pending creation"
}