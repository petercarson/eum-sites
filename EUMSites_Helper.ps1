function LoadEnvironmentSettings()
{
	[xml]$config = Get-Content -Path "$DistributionFolder\sharepoint.config"

    $environmentId = $config.settings.common.defaultEnvironment

    if ($environmentId -eq "") {
	    #-----------------------------------------------------------------------
	    # Prompt for the environment defined in the config
	    #-----------------------------------------------------------------------

        Write-Host "`n***** AVAILABLE ENVIRONMENTS *****" -ForegroundColor DarkGray
        $config.settings.environments.environment | ForEach {
            Write-Host "$($_.id)`t $($_.name) - $($_.webApp.adminSiteURL)"
        }
        Write-Host "***** AVAILABLE ENVIRONMENTS *****" -ForegroundColor DarkGray

        Do
        {
            [int]$environmentId = Read-Host "Enter the ID of the environment from the above list"
        }
        Until ($environmentId -gt 0)
    }

    [System.Xml.XmlLinkedNode]$Global:environment = $config.settings.environments.environment | Where { $_.id -eq $environmentId }

    # Set variables based on environment selected
    [string]$Global:WebAppURL = $environment.webApp.url
    [string]$Global:TenantAdminURL = $environment.webApp.adminSiteURL
    [string]$Global:SitesListSiteURL = "$($WebAppURL)$($environment.webApp.sitesListSiteCollectionPath)"
    [string]$Global:SiteListName = $config.settings.common.siteLists.siteListName
    [string]$Global:ManagedCredentials = $environment.webApp.managedCredentials
    [string]$Global:ManagedCredentialsType = $environment.webApp.managedCredentialsType

    [string]$Global:EUMClientID = $environment.EUM.clientID
    [string]$Global:EUMSecret = $environment.EUM.secret
    [string]$Global:Domain_FK = $environment.EUM.domain_FK
    [string]$Global:SystemConfiguration_FK = $environment.EUM.systemConfiguration_FK
    [string]$Global:EUMURL = $environment.EUM.EUMURL

    Write-Host "Environment set to $($environment.name) - $($environment.webApp.adminSiteURL) `n" -ForegroundColor Cyan

	#-----------------------------------------------------------------------
	# Get credentials from Windows Credential Manager
	#-----------------------------------------------------------------------
	if (Get-InstalledModule -Name "CredentialManager" -RequiredVersion "2.0") 
	{
		$Global:credentials = Get-StoredCredential -Target $managedCredentials 
        if ($managedCredentialsType -eq "UsernamePassword") {
		    if ($credentials -eq $null) {
			    $UserName = Read-Host "Enter the username to connect with"
			    $Password = Read-Host "Enter the password for $UserName" -AsSecureString 
			    $SaveCredentials = Read-Host "Save the credentials in Windows Credential Manager (Y/N)?"
			    if (($SaveCredentials -eq "y") -or ($SaveCredentials -eq "Y")) {
				    $temp = New-StoredCredential -Target $managedCredentials -UserName $UserName -SecurePassword $Password
			    }
			    $Global:SPCredentials = New-Object -typename System.Management.Automation.PSCredential -argumentlist $UserName,$Password
		    }
		    else {
			    $Global:SPCredentials = New-Object -typename System.Management.Automation.PSCredential -argumentlist $credentials.UserName,$credentials.Password
                Write-Host "Connecting with username" $credentials.UserName
		    }
        }
        else
        {
		    if ($credentials -eq $null) {
                [string]$Global:AppClientID = Read-Host "Enter the Client Id to connect with"
                [string]$Global:AppClientSecret = Read-Host "Enter the Secret"
			    $SaveCredentials = Read-Host "Save the credentials in Windows Credential Manager (Y/N)?"
			    if (($SaveCredentials -eq "y") -or ($SaveCredentials -eq "Y")) {
				    $temp = New-StoredCredential -Target $managedCredentials -UserName $AppClientID -Password $AppClientSecret
			    }
		    }
		    else {
                [string]$Global:AppClientID = $credentials.UserName
                [string]$Global:AppClientSecret = $credentials.GetNetworkCredential().password
                Write-Host "Connecting with Client Id" $AppClientID
		    }
        }
	}
	else
	{
		Write-Host "Required Windows Credential Manager 2.0 PowerShell Module not found. Please install the module by entering the following command in PowerShell: ""Install-Module -Name ""CredentialManager"" -RequiredVersion 2.0"""
		break
	}
}

function Helper-Connect-PnPOnline()
{
    Param
    (
        [Parameter(Mandatory=$true)][string] $URL
    )

    if (($AppClientID -ne "") -and ($AppClientSecret -ne "")) {
        Connect-PnPOnline -Url $URL -AppId $AppClientID -AppSecret $AppClientSecret
        }
    else {
        Connect-PnPOnline -Url $URL -Credentials $credentials
        }
}

function CreateSites()
{
    Param
    (
        [Parameter(Mandatory=$false)] $listItemID
    )

    if ($listItemID -ne $null)
    {
        # Get the specific Site Collection List item in master site for the site that needs to be created
        Helper-Connect-PnPOnline -Url $SitesListSiteURL

        $pendingSiteCollections = Get-PnPListItem -List $SiteListName -Query "
        <View>
            <Query>
                <Where>
                    <Eq>
                        <FieldRef Name='ID'/>
                        <Value Type='Integer'>$itemId</Value>
                    </Eq>
                </Where>
            </Query>
            <ViewFields>
                <FieldRef Name='ID'></FieldRef>
                <FieldRef Name='Title'></FieldRef>
                <FieldRef Name='EUMSiteURL'></FieldRef>
                <FieldRef Name='EUMAlias'></FieldRef>
                <FieldRef Name='EUMPublicGroup'></FieldRef>
                <FieldRef Name='EUMSetComposedLook'></FieldRef>
                <FieldRef Name='EUMBrandingDeploymentType'></FieldRef>
                <FieldRef Name='EUMBreadcrumbHTML'></FieldRef>
                <FieldRef Name='EUMParentURL'></FieldRef>
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
                <FieldRef Name='EUMSetComposedLook'></FieldRef>
                <FieldRef Name='EUMBrandingDeploymentType'></FieldRef>
                <FieldRef Name='EUMBreadcrumbHTML'></FieldRef>
                <FieldRef Name='EUMParentURL'></FieldRef>
                <FieldRef Name='EUMSiteTemplate'></FieldRef>
            </ViewFields>
        </View>"
    }

    if ($pendingSiteCollections.Count -gt 0)
    {
        # Get the time zone of the master site
        $spWeb = Get-PnPWeb -Includes RegionalSettings.TimeZone
        [int]$timeZoneId = $spWeb.RegionalSettings.TimeZone.Id

        # Iterate through the pending sites. Create them if needed, and apply template
        $pendingSiteCollections | ForEach {
            $pendingSite = $_

            [string]$siteTitle = $pendingSite["Title"]
            [string]$alias = $pendingSite["EUMAlias"]
            if ($alias)
            {
                $siteURL = "$($WebAppURL)/sites/$alias"
            }
            else
            {
                [string]$siteURL = ($pendingSite["EUMSiteURL"]).Url
            }
            [string]$publicGroup = $pendingSite["EUMPublicGroup"]
            [string]$breadcrumbHTML = $pendingSite["EUMBreadcrumbHTML"]
            [string]$parentURL = $pendingSite["EUMParentURL"].Url

            [bool]$siteCollection = CheckIfSiteCollection -siteURL $siteURL

            [string]$eumSiteTemplate = $pendingSite["EUMSiteTemplate"]

            $baseSiteTemplate = ""
            $baseSiteType = ""
            $pnpSiteTemplate = ""
            $siteCreated = $false

            switch ($eumSiteTemplate)
            {
                "Classic Team Site"
                    {
                    $baseSiteTemplate = "STS#0"
                    $baseSiteType = ""
                    }
                "Modern Communication Site"
                    {
                    $baseSiteTemplate = ""
                    $baseSiteType = "CommunicationSite"
                    }
                "Modern Team Site"
                    {
                    $baseSiteTemplate = ""
                    $baseSiteType = "TeamSite"
                    }
                "Modern Client Site"
                    {
                    $baseSiteTemplate = ""
                    $baseSiteType = "TeamSite"
                    $pnpSiteTemplate = $DistributionFolder + "\SiteTemplates\Client-Template-Template.xml"
                    }
            }

            # Classic style sites
            if ($baseSiteTemplate)
            {
                # Create the site
                if ($siteCollection)
                {
                    # Create site (if it exists, it will error but not modify the existing site)
                    Write-Host "Creating site collection $($siteURL) with base template $($baseSiteTemplate). Please wait..." -ForegroundColor Yellow
                    New-PnPTenantSite -Title $siteTitle -Url $siteURL -Owner $AdminUserName -TimeZone $timeZoneId -Template $baseSiteTemplate -RemoveDeletedSite -Wait -Force
                }
                else
                {
                    # Connect to parent site
                    Helper-Connect-PnPOnline -Url $parentURL

                    # Create the subsite
                    Write-Host "Creating subsite $($siteURL) with base template $($baseSiteTemplate) under $($parentURL). Please wait..." -ForegroundColor Yellow

                    [string]$subsiteURL = $siteURL.Replace($parentURL, "").Trim('/')
                    New-PnPWeb -Title $siteTitle -Url $subsiteURL -Template $baseSiteTemplate

                    Disconnect-PnPOnline
                }
                $siteCreated = $true

            }
            # Modern style sites
            else
            {
                # Create the site
                Write-Host "Creating site collection $($siteURL) with modern type $($baseSiteType). Please wait..." -ForegroundColor Yellow
                switch ($baseSiteType)
                {
                    "CommunicationSite"
                    {
                        New-PnPSite -Type CommunicationSite -Title $siteTitle -Url $siteURL
                        $siteCreated = $true
                    }
                    "TeamSite"
                    {
                        if ($publicGroup)
                        {
                            New-PnPSite -Type TeamSite -Title $siteTitle -Alias $alias -IsPublic
                        }
                        else
                        {
                            New-PnPSite -Type TeamSite -Title $siteTitle -Alias $alias
                        }
                        $siteCreated = $true
                    }
                }

            }

            if ($siteCreated)
            {
                if ($pnpSiteTemplate)
                {
                    Helper-Connect-PnPOnline -Url $siteURL
                    Apply-PnPProvisioningTemplate -Path $pnpSiteTemplate -ExcludeHandlers Publishing, ComposedLook, Navigation
                    Disconnect-PnPOnline
                }
            
                # Reconnect to the master site and update the site collection list
                Helper-Connect-PnPOnline -Url $SitesListSiteURL

                # Set the breadcrumb HTML
                [string]$siteRelativeURL = $siteURL.Replace($($WebAppURL), "")
                [string]$parentRelativeURL = $parentURL.Replace($($WebAppURL), "")
                $parentBreadcrumbHTML = ""
                if ($parentRelativeURL)
                {
                    $parentListItem = GetSiteEntry -siteRelativeURL $parentRelativeURL
                    if ($parentListItem)
                    {
                        [string]$parentBreadcrumbHTML = $parentListItem["EUMBreadcrumbHTML"]
                    }
                }
                [string]$breadcrumbHTML = GetBreadcrumbHTML -siteRelativeURL $SiteRelativeURL -siteTitle $siteTitle -parentBreadcrumbHTML $parentBreadcrumbHTML

                # Set the site created date, breadcrumb, and site URL
                [Microsoft.SharePoint.Client.ListItem]$spListItem = Set-PnPListItem -List $SiteListName -Identity $pendingSite.Id -Values @{ "EUMSiteCreated" = [System.DateTime]::Now; "EUMBreadcrumbHTML" = $breadcrumbHTML; "EUMSiteURL" = $siteRelativeURL }
            }

            # Reconnect to the master site for the next iteration
            Helper-Connect-PnPOnline -Url $SitesListSiteURL
        }
    }
    else
    {
        Write-Host "No sites pending creation" -ForegroundColor Green
    }
}

function CheckIfSiteCollection()
{
    Param
    (
        [Parameter(Mandatory=$true)][string] $siteURL
    )
    [bool] $isSiteCollection = $false
    foreach($managedPath in $managedPaths)
    {
        
        [string]$relativeURL = $siteURL.Replace($WebAppURL, "").ToLower().Trim()

        if ($relativeURL -eq '/')
        {
            $isSiteCollection = $true
        }
        elseif ($relativeURL.StartsWith(($managedPath.ToLower())))
        {
            [string]$relativeURLUpdated = $relativeURL.Replace($managedPath.ToLower(), "").Trim('/')
            [int]$charCount = ($relativeURLUpdated.ToCharArray() | Where-Object {$_ -eq '/'} | Measure-Object).Count
            
            $isSiteCollection = $charCount -eq 0
        }
    }

    return $isSiteCollection
}

function CheckIfSiteExists()
{
    Param
    (
        [Parameter(Mandatory=$true)][string] $siteURL,
        [Parameter(Mandatory=$false)][switch] $disconnect
    )

    try
    {
        Connect-PnPOnline -Url $siteURL -Credentials $credentials

        if ($disconnect.IsPresent)
        {
            Disconnect-PnPOnline
        }

        return $true
    }
    catch [System.Net.WebException]
    {
        if ([int]$_.Exception.Response.StatusCode -eq 404)
        {
            return $false
        }
        else
        {
            Write-Host "Exception Type: $($_.Exception.GetType().FullName)" -ForegroundColor Red
            Write-Host "Exception Message: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    catch
    {
        Write-Host "Exception Type: $($_.Exception.GetType().FullName)" -ForegroundColor Red
        Write-Host "Exception Message: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function GetParentWebURL()
{
    Param
    (
        [Parameter(Mandatory=$true)][string] $siteURL,
        [Parameter(Mandatory=$false)][switch] $disconnect
    )

    Connect-PnPOnline -Url $siteURL -Credentials $credentials
    [Microsoft.SharePoint.Client.Web]$spWeb = Get-PnPWeb -Includes ParentWeb.ServerRelativeUrl

    if ($disconnect.IsPresent)
    {
        Disconnect-PnPOnline
    }

    return $spWeb.ParentWeb.ServerRelativeUrl
}

function GetSubWebs()
{
    Param
    (
        [Parameter(Mandatory=$true)][string] $siteURL,
        [Parameter(Mandatory=$false)][switch] $disconnect
    )
    
    Connect-PnPOnline -Url $siteURL -Credentials $credentials
    [Microsoft.SharePoint.Client.Web]$spWeb = Get-PnPWeb -Includes Webs

    if ($spWeb.Webs.Count -gt 0)
    {
        $spSubWebs = Get-PnPSubWebs -Web $spWeb -Recurse
    }
    else
    {
        $spSubWebs = $null
    }

    if ($disconnect.IsPresent)
    {
        Disconnect-PnPOnline
    }

    return $spSubWebs
}

function GetBreadcrumbHTML()
{
    Param
    (
        [Parameter(Mandatory=$true)][string] $siteRelativeURL,
        [Parameter(Mandatory=$true)][string] $siteTitle,
        [Parameter(Mandatory=$false)][string] $parentBreadcrumbHTML
    )
    [string]$breadcrumbHTML = "<a href=`"$($siteRelativeURL)`">$($siteTitle)</a>"
	if ($parentBreadcrumbHTML)
	{
		$breadcrumbHTML = $parentBreadcrumbHTML + ' &gt; ' + $breadcrumbHTML
	}
    return $breadcrumbHTML
}

function PrepareSiteItemValues()
{
    Param
    (
        [parameter(Mandatory=$true)][string]$siteRelativeURL,
        [parameter(Mandatory=$true)][string]$siteTitle,
        [parameter(Mandatory=$false)]$parentURL,
        [parameter(Mandatory=$false)][string]$breadcrumbHTML,
        [parameter(Mandatory=$false)]$brandingDeploymentType,
        [parameter(Mandatory=$false)]$selectedThemeName,
        [parameter(Mandatory=$false)]$masterPageName,
        [parameter(Mandatory=$false)]$siteTemplateName,
        [parameter(Mandatory=$false)]$siteCreatedDate,
        [parameter(Mandatory=$false)]$subSite
    )

    [hashtable]$newListItemValues = @{}
    $newListItemValues.Add("Title", $siteTitle)
    $newListItemValues.Add("EUMSiteURL", $siteRelativeURL)
    
    if ($parentURL)
    {
        $newListItemValues.Add("EUMParentURL", $parentURL)
    }

    if ($breadcrumbHTML)
    {
        $newListItemValues.Add("EUMBreadcrumbHTML", $breadcrumbHTML)
    }

    if ($brandingDeploymentType)
    {
        $newListItemValues.Add("EUMBrandingDeploymentType", $brandingDeploymentType)
    }
    if ($selectedThemeName) 
    {
        $newListItemValues.Add("EUMSetComposedLook", $selectedThemeName)
    }
    if ($masterPageName)
    {
        $newListItemValues.Add("EUMSetMasterPage", $masterPageName)
    }
    if ($siteTemplateName)
    { 
        $newListItemValues.Add("EUMSiteTemplate", $siteTemplateName)
    }
    if ($siteCreatedDate)
    {
        $newListItemValues.Add("EUMSiteCreated", $siteCreatedDate)
    }

    $newListItemValues.Add("EUMIsSubsite", $subSite)

    return $newListItemValues
}

function GetSiteEntry()
{
    Param
    (
        [parameter(Mandatory=$true)][string]$siteRelativeURL,
        [Parameter(Mandatory=$false)][switch] $disconnect
    )
    
    Connect-PnPOnline -Url $SitesListSiteURL -Credentials $SPCredentials

    $siteListItem = Get-PnPListItem -List $SiteListName -Query "
    <View>
        <Query>
            <Where>
                <Eq>
                    <FieldRef Name='EUMSiteURL'/>
                    <Value Type='URL'>$($siteRelativeURL)</Value>
                </Eq>
            </Where>
        </Query>
    </View>"
    
    if ($disconnect.IsPresent)
    {
        Disconnect-PnPOnline
    }

    return $siteListItem
}

function AddOrUpdateSiteEntry()
{
    Param
    (
        [parameter(Mandatory=$true)][string]$siteRelativeURL,
        [parameter(Mandatory=$true)][string]$siteTitle,
        [parameter(Mandatory=$false)]$parentURL,
        [parameter(Mandatory=$false)][string]$breadcrumbHTML,
        [parameter(Mandatory=$false)][string]$brandingDeploymentType,
        [parameter(Mandatory=$false)]$selectedTheme,
        [parameter(Mandatory=$false)]$siteTemplateName,
        [parameter(Mandatory=$false)]$siteCreatedDate,
        [parameter(Mandatory=$false)]$spSubWebs
    )

    Write-Host "Adding $($siteTitle) to the $($SiteListName) list. Please wait..." -ForegroundColor Yellow

    [hashtable]$newListItemValues = PrepareSiteItemValues -siteRelativeURL $siteRelativeURL -siteTitle $siteTitle -parentURL $parentURL `
        -breadcrumbHTML $breadcrumbHTML -brandingDeploymentType $brandingDeploymentType -selectedThemeName $selectedTheme.name `
        -masterPageName $selectedTheme.masterPage -siteTemplateName $siteTemplateName -siteCreatedDate $siteCreatedDate


    $existingItem = GetSiteEntry -siteRelativeURL $siteRelativeURL -disconnect

    Connect-PnPOnline -Url $SitesListSiteURL -Credentials $SPCredentials
    if ($existingItem)
    {
        Write-Host "$($siteTitle) exists in $($SiteListName) list. Updating..." -ForegroundColor Yellow
        [Microsoft.SharePoint.Client.ListItem]$newListItem = Set-PnPListItem -Identity $existingItem.Id -List $SiteListName -Values $newListItemValues -ContentType "EUM Site Collection List"
    }
    else
    {
        [Microsoft.SharePoint.Client.ListItem]$newListItem = Add-PnPListItem -List $SiteListName -Values $newListItemValues -ContentType "EUM Site Collection List"
    }

    if ($newListItem)
    {
        Write-Host "The site $($siteTitle) was added to the $($SiteListName) list successfully" -ForegroundColor Green
    }

    # -----------
    # Subsites
    # -----------
    if ($spSubWebs)
    {
        Write-Host "Adding subsites of $($siteTitle) to $($SiteListName). Please wait..." -ForegroundColor Yellow
        foreach ($spSubWeb in $spSubWebs)
        {
            [string]$siteRelativeURL = $spSubWeb.ServerRelativeUrl
            [string]$siteTitle = $spSubWeb.Title
            $siteCreatedDate = $spSubWeb.Created.Date

            [string]$parentURL = GetParentWebURL -siteURL "$($WebAppURL)$($siteRelativeURL)" -disconnect

            if ($parentURL)
            {
                $parentListItem = GetSiteEntry -siteRelativeURL $parentURL -disconnect
                
                if ($parentListItem)
                {
                    [string]$parentBreadcrumbHTML = $parentListItem["EUMBreadcrumbHTML"]
                }
            }

            [string]$breadcrumbHTML = GetBreadcrumbHTML -siteRelativeURL $SiteRelativeURL -siteTitle $siteTitle -parentBreadcrumbHTML $parentBreadcrumbHTML

            [hashtable]$newListItemValues = PrepareSiteItemValues -siteRelativeURL $siteRelativeURL -siteTitle $siteTitle -parentURL $parentURL `
                -breadcrumbHTML $breadcrumbHTML -brandingDeploymentType $brandingDeploymentType -selectedThemeName $selectedTheme.name `
                -masterPageName (Split-Path $spSubWeb.CustomMasterUrl -Leaf) -siteTemplateName $siteTemplateName -siteCreatedDate $siteCreatedDate

            $existingItem = GetSiteEntry -siteRelativeURL $siteRelativeURL -disconnect

            Connect-PnPOnline -Url $SitesListSiteURL -Credentials $SPCredentials
            if ($existingItem)
            {
                Write-Host "$($siteTitle) exists in $($SiteListName) list. Updating..." -ForegroundColor Yellow
                [Microsoft.SharePoint.Client.ListItem]$newListItem = Set-PnPListItem -Identity $existingItem.Id -List $SiteListName -Values $newListItemValues -ContentType "EUM Site Collection List"
            }
            else
            {
                [Microsoft.SharePoint.Client.ListItem]$newListItem = Add-PnPListItem -List $SiteListName -Values $newListItemValues -ContentType "EUM Site Collection List"
            }

            if ($newListItem)
            {
                Write-Host "The site $($siteTitle) was added to the $($SiteListName) list successfully" -ForegroundColor Green
            }
        }
    }
}

function AddSiteEntry()
{
    Param
    (
        [parameter(Mandatory=$true)][string]$siteRelativeURL,
        [parameter(Mandatory=$true)][string]$siteTitle,
        [parameter(Mandatory=$false)]$parentURL,
        [parameter(Mandatory=$false)][string]$breadcrumbHTML,
        [parameter(Mandatory=$false)][string]$brandingDeploymentType,
        [parameter(Mandatory=$false)]$selectedTheme,
        [parameter(Mandatory=$false)]$siteTemplateName,
        [parameter(Mandatory=$false)]$siteCreatedDate,
        [parameter(Mandatory=$false)]$spSubWebs
    )

    $existingItem = GetSiteEntry -siteRelativeURL $siteRelativeURL -disconnect

    if (!$existingItem)
    {
        Connect-PnPOnline -Url $SitesListSiteURL -Credentials $SPCredentials

        Write-Host "Adding $($siteTitle) to the $($SiteListName) list. Please wait..." -ForegroundColor Yellow

        [hashtable]$newListItemValues = PrepareSiteItemValues -siteRelativeURL $siteRelativeURL -siteTitle $siteTitle -parentURL $parentURL `
            -breadcrumbHTML $breadcrumbHTML -brandingDeploymentType $brandingDeploymentType -selectedThemeName $selectedTheme.name `
            -masterPageName $selectedTheme.masterPage -siteTemplateName $siteTemplateName -siteCreatedDate $siteCreatedDate -subSite $false
        [Microsoft.SharePoint.Client.ListItem]$newListItem = Add-PnPListItem -List $SiteListName -Values $newListItemValues -ContentType "EUM Site Collection List"
    

        if ($newListItem)
        {
            Write-Host "The site $($siteTitle) was added to the $($SiteListName) list successfully" -ForegroundColor Green
        }

        # -----------
        # Subsites
        # -----------
        if ($spSubWebs)
        {
            Write-Host "Adding subsites of $($siteTitle) to $($SiteListName). Please wait..." -ForegroundColor Yellow
            foreach ($spSubWeb in $spSubWebs)
            {
                [string]$siteRelativeURL = $spSubWeb.ServerRelativeUrl
                [string]$siteTitle = $spSubWeb.Title
                $siteCreatedDate = $spSubWeb.Created.Date

                [string]$parentURL = GetParentWebURL -siteURL "$($WebAppURL)$($siteRelativeURL)" -disconnect

                if ($parentURL)
                {
                    $parentListItem = GetSiteEntry -siteRelativeURL $parentURL -disconnect
                
                    if ($parentListItem)
                    {
                        [string]$parentBreadcrumbHTML = $parentListItem["EUMBreadcrumbHTML"]
                    }
                }

                [string]$breadcrumbHTML = GetBreadcrumbHTML -siteRelativeURL $SiteRelativeURL -siteTitle $siteTitle -parentBreadcrumbHTML $parentBreadcrumbHTML

                [hashtable]$newListItemValues = PrepareSiteItemValues -siteRelativeURL $siteRelativeURL -siteTitle $siteTitle -parentURL $parentURL `
                    -breadcrumbHTML $breadcrumbHTML -brandingDeploymentType $brandingDeploymentType -selectedThemeName $selectedTheme.name `
                    -masterPageName (Split-Path $spSubWeb.CustomMasterUrl -Leaf) -siteTemplateName $siteTemplateName -siteCreatedDate $siteCreatedDate -subSite $true

                $existingItem = GetSiteEntry -siteRelativeURL $siteRelativeURL -disconnect

                Connect-PnPOnline -Url $SitesListSiteURL -Credentials $SPCredentials
                if ($existingItem)
                {
                    Write-Host "$($siteTitle) exists in $($SiteListName) list. Updating..." -ForegroundColor Yellow
                    [Microsoft.SharePoint.Client.ListItem]$newListItem = Set-PnPListItem -Identity $existingItem.Id -List $SiteListName -Values $newListItemValues -ContentType "EUM Site Collection List"
                }
                else
                {
                    [Microsoft.SharePoint.Client.ListItem]$newListItem = Add-PnPListItem -List $SiteListName -Values $newListItemValues -ContentType "EUM Site Collection List"
                }

                if ($newListItem)
                {
                    Write-Host "The site $($siteTitle) was added to the $($SiteListName) list successfully" -ForegroundColor Green
                }
            }
        }
    }
	else
	{
		Write-Host "The site $($siteTitle) exists in $($SiteListName) list. Skipping..." -ForegroundColor Yellow
	}
}

