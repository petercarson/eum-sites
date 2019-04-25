function LoadEnvironmentSettings() {

    [string]$Global:pnpTemplatePath = "c:\pnptemplates"

    # Check if running in Azure Automation or locally
    $Global:AzureAutomation = (Get-Command "Get-AutomationVariable" -errorAction SilentlyContinue)
    if ($AzureAutomation) {
        # Get automation variables
        $Global:SPCredentials = Get-AutomationPSCredential -Name 'SPOnlineCredentials'

        [string]$Global:SiteListName = Get-AutomationVariable -Name 'SiteListName'
        [string]$Global:WebAppURL = Get-AutomationVariable -Name 'WebAppURL'
        [string]$Global:AdminURL = $WebAppURL.Replace(".sharepoint.com", "-admin.sharepoint.com")
        [string]$Global:SitesListSiteURL = "$($WebAppURL)$(Get-AutomationVariable -Name 'SitesListSiteURL')"

        if ($loadEUMCredentials) {
            # $Global:EUMClientID = $ManagedEUMCredentials.UserName
            # $Global:EUMSecret = (New-Object PSCredential "user", $ManagedEUMCredentials.Password).GetNetworkCredential().Password
        }

        if ($loadGraphAPICredentials) {
            $Global:AADCredentials = Get-AutomationPSCredential -Name 'AADCredentials'
            $Global:AADClientID = $AADCredentials.UserName
            $Global:AADSecret = (New-Object PSCredential "user", $AADCredentials.Password).GetNetworkCredential().Password
            $Global:AADDomain = Get-AutomationVariable -Name 'AADDomain'
        }
    }
    else {
        [xml]$config = Get-Content -Path "$DistributionFolder\sharepoint.config"

        [System.Array]$Global:managedPaths = $config.settings.common.managedPaths.path
        [string]$Global:SiteListName = $config.settings.common.siteLists.siteListName

        $environmentId = $config.settings.common.defaultEnvironment

        if (-not $environmentId) {
            #-----------------------------------------------------------------------
            # Prompt for the environment defined in the config
            #-----------------------------------------------------------------------

            Write-Host "`n***** AVAILABLE ENVIRONMENTS *****" -ForegroundColor DarkGray
            $config.settings.environments.environment | ForEach {
                Write-Host "$($_.id)`t $($_.name) - $($_.webApp.URL)"
            }
            Write-Host "***** AVAILABLE ENVIRONMENTS *****"

            Do {
                [int]$environmentId = Read-Host "Enter the ID of the environment from the above list"
            }
            Until ($environmentId -gt 0)
        }

        [System.Xml.XmlLinkedNode]$Global:environment = $config.settings.environments.environment | Where { $_.id -eq $environmentId }

        # Set variables based on environment selected
        [string]$Global:WebAppURL = $environment.webApp.url
        [string]$Global:AdminURL = $environment.webApp.url.Replace(".sharepoint.com", "-admin.sharepoint.com")
        [string]$Global:SitesListSiteURL = "$($WebAppURL)$($environment.webApp.sitesListSiteCollectionPath)"
        
        Write-Host "Environment set to $($environment.name) - $($environment.webApp.URL) `n"

        $Global:SPCredentials = GetManagedCredentials -managedCredentials $environment.webApp.managedCredentials -ManagedCredentialsType $environment.webApp.managedCredentialsType

        if ($loadEUMCredentials) {
            $ManagedEUMCredentials = GetManagedCredentials -managedCredentials $environment.EUM.managedCredentials -ManagedCredentialsType $environment.EUM.managedCredentialsType
            $Global:EUMClientID = $ManagedEUMCredentials.UserName
            $Global:EUMSecret = (New-Object PSCredential "user", $ManagedEUMCredentials.Password).GetNetworkCredential().Password
        }

        if ($loadGraphAPICredentials) {
            $AADCredentials = GetManagedCredentials -managedCredentials $environment.graphAPI.managedCredentials -ManagedCredentialsType $environment.graphAPI.managedCredentialsType
            $Global:AADClientID = $AADCredentials.UserName
            $Global:AADSecret = (New-Object PSCredential "user", $AADCredentials.Password).GetNetworkCredential().Password
            $Global:AADDomain = $environment.graphAPI.AADDomain
        }
    }
}

function GetManagedCredentials()
{
    [OutputType([System.Management.Automation.PSCredential])]
    Param
    (
        [Parameter(Mandatory=$true)][string] $managedCredentials,
        [Parameter(Mandatory=$true)][string] $managedCredentialsType
    )

    if (-not(Get-InstalledModule -Name "CredentialManager" -RequiredVersion "2.0")) {
        Write-Host "Required Windows Credential Manager 2.0 PowerShell Module not found. Please install the module by entering the following command in PowerShell: ""Install-Module -Name ""CredentialManager"" -RequiredVersion 2.0"""
        return $null
    }

    #-----------------------------------------------------------------------
    # Get credentials from Windows Credential Manager
    #-----------------------------------------------------------------------
    $Credentials = Get-StoredCredential -Target $managedCredentials 
    switch ($managedCredentialsType) {
        "UsernamePassword" {
            if ($Credentials -eq $null) {
                $UserName = Read-Host "Enter the username to connect with for $managedCredentials"
                $Password = Read-Host "Enter the password for $UserName" -AsSecureString 
                $SaveCredentials = Read-Host "Save the credentials in Windows Credential Manager (Y/N)?"
                if (($SaveCredentials -eq "y") -or ($SaveCredentials -eq "Y")) {
                    $temp = New-StoredCredential -Target $managedCredentials -UserName $UserName -SecurePassword $Password
                }
                $Credentials = New-Object -typename System.Management.Automation.PSCredential -argumentlist $UserName, $Password
            }
            else {
                Write-Host "Connecting with username" $Credentials.UserName
            }
        }

        "ClientIdSecret" {
            if ($Credentials -eq $null) {
                $ClientID = Read-Host "Enter the Client Id to connect with for $managedCredentials"
                $ClientSecret = Read-Host "Enter the Secret" -AsSecureString
                $SaveCredentials = Read-Host "Save the credentials in Windows Credential Manager (Y/N)?"
                if (($SaveCredentials -eq "y") -or ($SaveCredentials -eq "Y")) {
                    $temp = New-StoredCredential -Target $managedCredentials -UserName $ClientID -SecurePassword $ClientSecret
                }
                $Credentials = New-Object -typename System.Management.Automation.PSCredential -argumentlist $ClientID, $ClientSecret
            }
            else {
                Write-Host "Connecting with Client Id" $Credentials.UserName
            }
        }
    }

    return ($Credentials)
}

function Helper-Connect-PnPOnline()
{
    Param
    (
        [Parameter(Mandatory=$true)][string] $URL
    )

    if ($O365ClientID -and $O365ClientSecret) {
        Connect-PnPOnline -Url $URL -AppId $O365ClientID -AppSecret $O365ClientSecret
        }
    else {
        Connect-PnPOnline -Url $URL -Credentials $SPCredentials
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
        Helper-Connect-PnPOnline -Url $siteURL

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
            try {
                $spContext = New-Object Microsoft.SharePoint.Client.ClientContext($siteURL)
                $spContext.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($SPCredentials.UserName, $SPCredentials.Password)
                $web = $spContext.Web
                $spContext.Load($web)
                $spContext.ExecuteQuery()
            }
            catch
            {      
                if (($_.Exception.Message -like "*Cannot contact site at the specified URL*") -and ($_.Exception.Message -like "*There is no Web named*"))
                {
                    return $false
                }
            }
        }
    }
    catch
    {
        Write-Host "Exception Type: $($_.Exception.GetType().FullName)"
        Write-Host "Exception Message: $($_.Exception.Message)"
    }
}

function GetParentWebURL()
{
    Param
    (
        [Parameter(Mandatory=$true)][string] $siteURL,
        [Parameter(Mandatory=$false)][switch] $disconnect
    )

    Helper-Connect-PnPOnline -Url $siteURL
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
    
    Helper-Connect-PnPOnline -Url $siteURL
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
        [Parameter(Mandatory=$false)][string] $parentURL
    )
    [string]$parentBreadcrumbHTML = ""

    if ($parentURL)
    {
        $parentURL = $parentURL.Replace($WebAppURL, "")
        $parentListItem = GetSiteEntry -siteRelativeURL $parentURL
        if ($parentListItem)
        {
            [string]$parentBreadcrumbHTML = $parentListItem["EUMBreadcrumbHTML"]
        }
    }

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
        [parameter(Mandatory=$true)][string]$siteTitle,
        [parameter(Mandatory=$false)][string]$breadcrumbHTML,
        [parameter(Mandatory=$false)]$siteCreatedDate
    )

    [hashtable]$newListItemValues = @{}
    $newListItemValues.Add("Title", $siteTitle)

    if ($breadcrumbHTML)
    {
        $newListItemValues.Add("EUMBreadcrumbHTML", $breadcrumbHTML)
    }

    if ($siteCreatedDate)
    {
        $newListItemValues.Add("EUMSiteCreated", $siteCreatedDate)
    }

    return $newListItemValues
}

function GetSiteEntry()
{
    Param
    (
        [parameter(Mandatory=$true)][string]$siteRelativeURL,
        [Parameter(Mandatory=$false)][switch] $disconnect
    )
    
    Helper-Connect-PnPOnline -Url $SitesListSiteURL

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
        [parameter(Mandatory=$false)]$siteCreatedDate,
        [parameter(Mandatory=$false)]$spSubWebs
    )

    $existingItem = GetSiteEntry -siteRelativeURL $siteRelativeURL -disconnect
    if ($existingItem.Count -gt 1)
    {
        Write-Host "Error: Multiple existing list entries found for the same URL"
        Write-Host $existingItem
        return
    }

    Helper-Connect-PnPOnline -Url $SitesListSiteURL
    if ($existingItem)
    {
        [string]$breadcrumbHTML = GetBreadcrumbHTML -siteRelativeURL $SiteRelativeURL -siteTitle $siteTitle -parentURL $existingItem["EUMParentURL"].Url
        [hashtable]$newListItemValues = PrepareSiteItemValues -siteTitle $siteTitle -breadcrumbHTML $breadcrumbHTML -siteCreatedDate $siteCreatedDate

        $updateRequired = $false

        foreach ($newListItemKey in $newListItemValues.Keys)
        {
            if ($newListItemKey -eq "EUMBreadcrumbHTML")
            {
                if ($existingItem[$newListItemKey] -notlike "*$($newListItemValues[$newListItemKey])*")
                {
                    $updateRequired = $true
                }
            }
            elseif ($newListItemKey -ne "EUMSiteCreated")
            {
                if ($existingItem[$newListItemKey] -ne $newListItemValues[$newListItemKey])
                {
                    $updateRequired = $true
                }
            }
        }

        if ($updateRequired)
        {
            Write-Host "$($siteTitle) exists in $($SiteListName) list. Updating..."
            [Microsoft.SharePoint.Client.ListItem]$newListItem = Set-PnPListItem -Identity $existingItem.Id -List $SiteListName -Values $newListItemValues
        }
        else
        {
            Write-Host "$($siteTitle) exists in $($SiteListName) list. No updates required."
        }
    }
    else
    {
        [string]$breadcrumbHTML = GetBreadcrumbHTML -siteRelativeURL $SiteRelativeURL -siteTitle $siteTitle -parentURL ""
        [hashtable]$newListItemValues = PrepareSiteItemValues -siteTitle $siteTitle -breadcrumbHTML $breadcrumbHTML -siteCreatedDate $siteCreatedDate
        [Microsoft.SharePoint.Client.ListItem]$newListItem = Add-PnPListItem -List $SiteListName -Values $newListItemValues -ContentType "Base Site Request"
        Write-Host "The site $($siteTitle) was added to the $($SiteListName) list successfully"
    }

    # -----------
    # Subsites
    # -----------
    if ($spSubWebs)
    {
        Write-Host "Checking subsites of $($siteTitle) to $($SiteListName). Please wait..."
        foreach ($spSubWeb in $spSubWebs)
        {
            [string]$siteRelativeURL = $spSubWeb.ServerRelativeUrl
            [string]$siteTitle = $spSubWeb.Title
            $siteCreatedDate = $spSubWeb.Created.Date
            [string]$breadcrumbHTML = GetBreadcrumbHTML -siteRelativeURL $SiteRelativeURL -siteTitle $siteTitle -parentURL $parentURL

            [hashtable]$newListItemValues = PrepareSiteItemValues -siteTitle $siteTitle -breadcrumbHTML $breadcrumbHTML -siteCreatedDate $siteCreatedDate

            $existingItem = GetSiteEntry -siteRelativeURL $siteRelativeURL -disconnect

            Helper-Connect-PnPOnline -Url $SitesListSiteURL
            if ($existingItem)
            {
                $updateRequired = $false

                foreach ($newListItemKey in $newListItemValues.Keys)
                {
                    if ($newListItemKey -eq "EUMBreadcrumbHTML")
                    {
                        if ($existingItem[$newListItemKey] -notlike "*$($newListItemValues[$newListItemKey])*")
                        {
                            $updateRequired = $true
                        }
                    }
                    elseif ($newListItemKey -ne "EUMSiteCreated")
                    {
                        if ($existingItem[$newListItemKey] -ne $newListItemValues[$newListItemKey])
                        {
                            $updateRequired = $true
                        }
                    }
                }

                if ($updateRequired)
                {
                    Write-Host "$($siteTitle) exists in $($SiteListName) list. Updating..."
                    [Microsoft.SharePoint.Client.ListItem]$newListItem = Set-PnPListItem -Identity $existingItem.Id -List $SiteListName -Values $newListItemValues
                }
                else
                {
                    Write-Host "$($siteTitle) exists in $($SiteListName) list. No updates required."
                }
            }
            else
            {
                [Microsoft.SharePoint.Client.ListItem]$newListItem = Add-PnPListItem -List $SiteListName -Values $newListItemValues -ContentType "Base Site Request"
                Write-Host "The site $($siteTitle) was added to the $($SiteListName) list successfully"
            }
        }
    }
}

function SetSiteLogo {
    Param
    (
        [Parameter(Position = 0, Mandatory = $true)][string] $siteURL,
        [Parameter(Position = 1, Mandatory = $true)][string] $logoRelativeURL,
        [Parameter(Position = 2, Mandatory = $false)][switch] $subsitesInherit
    )

    try {            
        $web = Get-PnPWeb
        Set-PnpWeb -Web $web.Id -SiteLogoUrl $logoRelativeURL


        if ($subsitesInherit.IsPresent) {
            Write-Host "Updating subsites logo"

            $subwebs = Get-PnPSubWebs -Recurse
            Foreach ($web in $subwebs) {
                Set-PnpWeb -Web $web.Id -SiteLogoUrl $logoRelativeURL
            }
        }
        else {
            Write-Host "Updating site logo"
            $web = Get-PnPWeb
            Set-PnpWeb -Web $web -SiteLogoUrl $logoRelativeURL
        }
          
    }
    catch {
        Write-Host "An exception occurred setting site logo in $siteURL"
        Write-Host "Exception Type: $($_.Exception.GetType().FullName)" -ForegroundColor Red
        Write-Host "Exception Message: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function ApplyModernSiteBranding {

    Param
    (
        [Parameter(Position = 0, Mandatory = $true)][string] $siteURL,
        [Parameter(Position = 4, Mandatory = $true)][string] $logoFile,
        [Parameter(Position = 4, Mandatory = $true)][string] $homePageImage
    )
        
    #$cred = Get-Credential
    #Connect-PnPOnline $siteURL -Credentials $PScredentials
    Helper-Connect-PnPOnline -Url $siteURL

    # Apply a custom theme to a Modern Site

    # First, upload the theme assets
    Write-Host "`nUploading branding files..." -foreground yellow

    Add-PnPFile -Path "$DistributionFolder\Branding\$logoFile" -Folder SiteAssets    
    Add-PnPFile -Path "$DistributionFolder\Branding\$homePageImage" -Folder SiteAssets      

    Write-Host "`nBranding files deployed`n" -foreground green

    # Second, apply the theme assets to the site
    $web = Get-PnPWeb
    $logo = $web.ServerRelativeUrl + "/Style Library/$logoFile"

    Write-Host "Setting site logo..." -foreground yellow

    SetSiteLogo -siteURL $siteURL -logoRelativeURL $logo #-subsitesInherit

    # We use OOTB CSOM operation for this
    #$web.ApplyTheme($palette, $font, $background, $true)
    $web.Update()
    # Set timeout as high as possible and execute
    $web.Context.RequestTimeout = [System.Threading.Timeout]::Infinite
    $web.Context.ExecuteQuery()  
}


function DisableDenyAddAndCustomizePages {
    Param
    (
        [Parameter(Position = 0, Mandatory = $true)][string] $siteURL
    )

    Helper-Connect-PnPOnline -URL $AdminURL
    
    $context = Get-PnPContext
    $site = Get-PnPTenantSite -Detailed -Url $siteURL
     
    $site.DenyAddAndCustomizePages = [Microsoft.Online.SharePoint.TenantAdministration.DenyAddAndCustomizePagesStatus]::Disabled
     
    $site.Update()
    $context.ExecuteQuery()
    $context.Dispose()

    $status = $null
    do {
        Write-Host "Waiting...   $status"
        Start-Sleep -Seconds 5
        $site = Get-PnPTenantSite -Detailed -Url $siteURL
        $status = $site.Status
    
    } while ($status -ne 'Active')

    Disconnect-PnPOnline
}

Set-PnPTraceLog -On -Level Debug
