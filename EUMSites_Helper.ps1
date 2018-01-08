function LoadEnvironmentSettings()
{
	[xml]$config = Get-Content -Path "$DistributionFolder\sharepoint.config"

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

  [System.Xml.XmlLinkedNode]$Global:environment = $config.settings.environments.environment | Where { $_.id -eq $environmentId }

  # Set variables based on environment selected
  [string]$Global:WebAppURL = $environment.webApp.url
  [string]$Global:TenantAdminURL = $environment.webApp.adminSiteURL
  [string]$Global:SitesListSiteURL = "$($WebAppURL)$($environment.webApp.sitesListSiteCollectionPath)"
  [string]$Global:SiteListName = $config.settings.common.siteLists.siteListName
  [string]$Global:ManagedCredentials = $environment.webApp.managedCredentials

  [string]$Global:EUMClientID = $environment.EUM.clientID
  [string]$Global:EUMSecret = $environment.EUM.secret
  [string]$Global:Domain_FK = $environment.EUM.domain_FK
  [string]$Global:SystemConfiguration_FK = $environment.EUM.systemConfiguration_FK
  [string]$Global:EUMURL = $environment.EUM.EUMURL

  Write-Host "Environment set to $($environment.name) - $($environment.webApp.adminSiteURL) `n" -ForegroundColor Cyan

	#-----------------------------------------------------------------------
	# SharePoint Deployment
	#-----------------------------------------------------------------------
	$CredentialManager = "true"
	if (Get-InstalledModule -Name "CredentialManager" -RequiredVersion "2.0") 
	{
		$Global:credentials = Get-StoredCredential -Target $managedCredentials 
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
		Write-Host "Required Windows Credential Manager 2.0 PowerShell Module not found. Please install the module by entering the following command in PowerShell: ""Install-Module -Name ""CredentialManager"" -RequiredVersion 2.0"""
		break
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

