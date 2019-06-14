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
        [string]$Global:SiteCollectionAdministrator = Get-AutomationVariable -Name 'siteCollectionAdministrator'

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
        [string]$Global:SiteCollectionAdministrator = $environment.webApp.siteCollectionAdministrator
        
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

function GetBreadcrumbHTML()
{
    Param
    (
        [Parameter(Mandatory=$true)][string] $siteURL,
        [Parameter(Mandatory=$true)][string] $siteTitle,
        [Parameter(Mandatory=$false)][string] $parentURL
    )
    [string]$parentBreadcrumbHTML = ""

    if ($parentURL)
    {
				Helper-Connect-PnPOnline -Url $SitesListSiteURL

				$parentListItem = Get-PnPListItem -List $SiteListName -Query "
				<View>
						<Query>
								<Where>
										<Eq>
												<FieldRef Name='EUMSiteURL'/>
												<Value Type='Text'>$($siteURL)</Value>
										</Eq>
								</Where>
						</Query>
				</View>"

        if ($parentListItem)
        {
            [string]$parentBreadcrumbHTML = $parentListItem["EUMBreadcrumbHTML"]
        }
        else
        {
						Write-Host "No entry found for $parentURL"
        }
    }

    $siteURL = $siteURL.Replace($webAppURL, "")
    [string]$breadcrumbHTML = "<a href=`"$($siteURL)`">$($siteTitle)</a>"
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
        [parameter(Mandatory=$false)][string]$siteURL,
        [parameter(Mandatory=$false)][string]$siteTitle,
        [parameter(Mandatory=$false)]$parentURL,
        [parameter(Mandatory=$false)][string]$breadcrumbHTML,
        [parameter(Mandatory=$false)]$siteCreatedDate,
        [parameter(Mandatory=$false)]$subSite
    )

    [hashtable]$newListItemValues = @{}

    if ($siteURL)
    {
        $newListItemValues.Add("EUMSiteURL", $siteURL)
    }

    if ($siteTitle)
    {
        $newListItemValues.Add("Title", $siteTitle)
    }

    if ($parentURL)
    {
        $newListItemValues.Add("EUMParentURL", $parentURL)
    }

    if ($breadcrumbHTML)
    {
        $newListItemValues.Add("EUMBreadcrumbHTML", $breadcrumbHTML)
    }

    if ($siteCreatedDate)
    {
        $newListItemValues.Add("EUMSiteCreated", $siteCreatedDate)
    }

    if ($subSite)
    {
        $newListItemValues.Add("EUMSubSite", $subSite)
    }

    return $newListItemValues
}

Set-PnPTraceLog -On -Level Debug
