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

# Get the config file
[xml]$config = Get-Content -Path "$DistributionFolder/sharepoint.config"

$hubSite = $config.settings.common.associatedHubSite.hubSiteUrl

if ($listItemID -gt 0) {
    # Get the specific Site Collection List item in master site for the site that needs to be created
    Helper-Connect-PnPOnline -Url $SitesListSiteURL

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
            <FieldRef Name='Author'></FieldRef>
        </ViewFields>
    </View>"
}
else {
    # Check the Site Collection List in master site for any sites that need to be created
    Helper-Connect-PnPOnline -Url $SitesListSiteURL

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
            [string]$siteURL = ($pendingSite["EUMSiteURL"]).Url
        }
        [string]$publicGroup = $pendingSite["EUMPublicGroup"]
        [string]$breadcrumbHTML = $pendingSite["EUMBreadcrumbHTML"]
        [string]$parentURL = $pendingSite["EUMParentURL"].Url
        [string]$Division = $pendingSite["EUMDivision"].LookupValue

        [bool]$siteCollection = CheckIfSiteCollection -siteURL $siteURL

        [string]$eumSiteTemplate = $pendingSite["EUMSiteTemplate"]

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
                $parentURL = $parentURL.Replace($WebAppURL, "")
            }
		}

        $baseSiteTemplate = ""
        $baseSiteType = ""
        $pnpSiteTemplate = ""
        $siteCreated = $false

        switch ($eumSiteTemplate) {
            "Modern Communication Site" {
                $baseSiteTemplate = ""
                $baseSiteType = "CommunicationSite"
            }
            "Modern Team Site" {
                $baseSiteTemplate = ""
                $baseSiteType = "TeamSite"
            }
            "Modern Client Site"
                {
                $baseSiteTemplate = ""
                $baseSiteType = "TeamSite"
                $pnpSiteTemplate = "$DistributionFolder\SiteTemplates\Client-Template-Template.xml"
            }
            "Client Communication Site"
                {
                $baseSiteTemplate = ""
                $baseSiteType = "CommunicationSite"
                $pnpSiteTemplate = "$DistributionFolder\SiteTemplates\Client-Template-Template.xml"
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

            # Set the site collection admins
            Add-PnPSiteCollectionAdmin -Owners "pcarson@envisionit.com"

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
            
            If (($eumSiteTemplate -eq "Modern Client Site") -or ($eumSiteTemplate -eq "Client Communication Site"))
            {
                Add-PnPFolder -Name "Quotes" -Folder "/Shared Documents"
                Add-PnPFolder -Name "Signed Quotes" -Folder "/Shared Documents/Quotes"
                Add-PnPFolder -Name "Invoices" -Folder "/Shared Documents"

                Add-PnPFolder -Name "Business Development" -Folder "/Private Documents"
                Add-PnPFolder -Name "Confidential" -Folder "/Private Documents"
                Add-PnPFolder -Name "Quotes" -Folder "/Private Documents"

                Remove-PnPContentTypeFromList -List "Shared Documents" -ContentType "Document"
                Remove-PnPContentTypeFromList -List "Private Documents" -ContentType "Document"

                Remove-PnPGroup -Identity "$siteTitle Members" -Force
                Remove-PnPGroup -Identity "$siteTitle Owners" -Force
                Remove-PnPGroup -Identity "$siteTitle Visitors" -Force
            }

            # Reconnect to the master site and update the site collection list
            Helper-Connect-PnPOnline -Url $SitesListSiteURL

            # Set the breadcrumb HTML
            [string]$breadcrumbHTML = GetBreadcrumbHTML -siteRelativeURL $siteURL -siteTitle $siteTitle -parentURL $parentURL

            # Set the site created date, breadcrumb, and site URL
            [Microsoft.SharePoint.Client.ListItem]$spListItem = Set-PnPListItem -List $SiteListName -Identity $pendingSite.Id -Values @{ "EUMSiteCreated" = [System.DateTime]::Now; "EUMBreadcrumbHTML" = $breadcrumbHTML; "EUMSiteURL" = $siteURL.Replace($WebAppURL, ""); "EUMParentURL" = $parentURL }

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