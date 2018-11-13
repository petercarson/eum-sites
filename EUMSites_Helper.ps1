function LoadEnvironmentSettings()
{
    [string]$Global:WebAppURL = $Env:webAppURL

    Set-PnPTraceLog -On -Level Debug

    if ($WebAppURL -ne "")
    {
        [string]$Global:SitesListSiteURL = "$($WebAppURL)$($Env:sitesListSiteCollectionPath)"
        [string]$Global:SiteListName = $Env:siteListName

        # MSI Variables via Function Application Settings Variables
        # Endpoint and Password
        $endpoint = $env:MSI_ENDPOINT
        $secret = $env:MSI_SECRET

        # Vault URI to get AuthN Token
        $vaultTokenURI = 'https://vault.azure.net&api-version=2017-09-01'
        # Create AuthN Header with our Function App Secret
        $header = @{'Secret' = $secret}

        # Get Key Vault AuthN Token
        $authenticationResult = Invoke-RestMethod -Method Get -Headers $header -Uri ($endpoint +'?resource=' +$vaultTokenURI)
        # Use Key Vault AuthN Token to create Request Header
        $requestHeader = @{ Authorization = "Bearer $($authenticationResult.access_token)" }

        # Our Key Vault Credential that we want to retreive URI
        # NOTE: API Ver for this is 2015-06-01

        # Call the Vault and Retrieve Creds
        $vaultSecretURI  = $Env:serviceAccountURI
        $Secret = Invoke-RestMethod -Method GET -Uri $vaultSecretURI -ContentType 'application/json' -Headers $requestHeader
        $UserName = $Secret.Value

        $vaultSecretURI = $Env:serviceAccountPasswordURI
        $Secret = Invoke-RestMethod -Method GET -Uri $vaultSecretURI -ContentType 'application/json' -Headers $requestHeader
        $Password = ConvertTo-SecureString $Secret.Value -AsPlainText -Force

        $Global:SPCredentials = New-Object -typename System.Management.Automation.PSCredential -argumentlist $UserName, $Password
    }
    else
    {
	    [xml]$config = Get-Content -Path "$DistributionFolder\sharepoint.config"

        [System.Array]$Global:managedPaths = $config.settings.common.managedPaths.path
        [string]$Global:SiteListName = $config.settings.common.siteLists.siteListName

        $environmentId = $config.settings.common.defaultEnvironment

        if (-not $environmentId) {
	        #-----------------------------------------------------------------------
	        # Prompt for the environment defined in the config
	        #-----------------------------------------------------------------------

            Write-Output "`n***** AVAILABLE ENVIRONMENTS *****" -ForegroundColor DarkGray
            $config.settings.environments.environment | ForEach {
                Write-Output "$($_.id)`t $($_.name) - $($_.webApp.URL)"
            }
            Write-Output "***** AVAILABLE ENVIRONMENTS *****"

            Do
            {
                [int]$environmentId = Read-Host "Enter the ID of the environment from the above list"
            }
            Until ($environmentId -gt 0)
        }

        [System.Xml.XmlLinkedNode]$Global:environment = $config.settings.environments.environment | Where { $_.id -eq $environmentId }

        # Set variables based on environment selected
        [string]$Global:WebAppURL = $environment.webApp.url
        [string]$Global:SitesListSiteURL = "$($WebAppURL)$($environment.webApp.sitesListSiteCollectionPath)"
        
        $ManagedCredentials = $environment.webApp.managedCredentials
        $ManagedCredentialsType = $environment.webApp.managedCredentialsType
        $TenantId = $environment.webApp.tenantId
        $VaultName = $environment.webApp.vaultName

        $ManagedEUMCredentials = $environment.EUM.managedEUMCredentials
        [string]$Global:Domain_FK = $environment.EUM.domain_FK
        [string]$Global:SystemConfiguration_FK = $environment.EUM.systemConfiguration_FK
        [string]$Global:EUMURL = $environment.EUM.EUMURL

        Write-Output "Environment set to $($environment.name) - $($environment.webApp.URL) `n"

	    #-----------------------------------------------------------------------
	    # Get credentials from Windows Credential Manager
	    #-----------------------------------------------------------------------
	    if (Get-InstalledModule -Name "CredentialManager" -RequiredVersion "2.0") 
	    {
		    $Global:SPCredentials = Get-StoredCredential -Target $managedCredentials 
            switch ($managedCredentialsType) 
            {
                "UsernamePassword"
                {
		            if ($SPCredentials -eq $null) {
			            $UserName = Read-Host "Enter the username to connect with"
			            $Password = Read-Host "Enter the password for $UserName" -AsSecureString 
			            $SaveCredentials = Read-Host "Save the credentials in Windows Credential Manager (Y/N)?"
			            if (($SaveCredentials -eq "y") -or ($SaveCredentials -eq "Y")) {
				            $temp = New-StoredCredential -Target $managedCredentials -UserName $UserName -SecurePassword $Password
			            }
			            $Global:SPCredentials = New-Object -typename System.Management.Automation.PSCredential -argumentlist $UserName,$Password
		            }
		            else {
                        Write-Output "Connecting with username" $SPCredentials.UserName
		            }
		        }

                "ClientIdSecret"
                {
		            if ($SPCredentials -eq $null) {
                        [string]$Global:O365ClientID = Read-Host "Enter the Client Id to connect with"
                        [string]$Global:O365ClientSecret = Read-Host "Enter the Secret"
			            $SaveCredentials = Read-Host "Save the credentials in Windows Credential Manager (Y/N)?"
			            if (($SaveCredentials -eq "y") -or ($SaveCredentials -eq "Y")) {
				            $temp = New-StoredCredential -Target $managedCredentials -UserName $O365ClientID -Password $O365ClientSecret
			            }
		            }
		            else {
                        [string]$Global:O365ClientID = $SPCredentials.UserName
                        [string]$Global:O365ClientSecret = $SPCredentials.GetNetworkCredential().password
                        Write-Output "Connecting with Client Id" $O365ClientID
		            }
                }

                "AzureKeyVault"
                {
		            if ($SPCredentials -eq $null) {
			            $UserName = Read-Host "Enter the username to connect to the Azure Key Vault with"
			            $Password = Read-Host "Enter the password for $UserName" -AsSecureString 
			            $SaveCredentials = Read-Host "Save the credentials in Windows Credential Manager (Y/N)?"
			            if (($SaveCredentials -eq "y") -or ($SaveCredentials -eq "Y")) {
				            $temp = New-StoredCredential -Target $managedCredentials -UserName $UserName -SecurePassword $Password
			            }
			            $Global:SPCredentials = New-Object -typename System.Management.Automation.PSCredential -argumentlist $UserName, $Password
		            }
		            else {
                        Write-Output "Connecting with username" $SPCredentials.UserName
		            }

		            Login-AzureRmAccount -Credential $SPCredentials -TenantId $TenantId
                    $UserName = (Get-AzureKeyVaultSecret -VaultName $VaultName -Name 'ServiceAccount').SecretValueText
                    $Password = ConvertTo-SecureString (Get-AzureKeyVaultSecret -VaultName $VaultName -Name 'ServiceAccountPassword').SecretValueText -AsPlainText -Force
			        $Global:SPCredentials = New-Object -typename System.Management.Automation.PSCredential -argumentlist $UserName, $Password
                }
            }
            if ($ManagedEUMCredentials -ne $null)
            {
    		    $EUMCredentials = Get-StoredCredential -Target $managedEUMCredentials
		        if ($EUMCredentials -eq $null) {
			        $Global:EUMClientID = Read-Host "Enter the Client ID to connect to EUM with"
			        $SecureEUMSecret = Read-Host "Enter the secret for $EUMClientID" -AsSecureString
			        $SaveCredentials = Read-Host "Save the credentials in Windows Credential Manager (Y/N)?"
			        if (($SaveCredentials -eq "y") -or ($SaveCredentials -eq "Y")) {
				        $temp = New-StoredCredential -Target $ManagedEUMCredentials -UserName $EUMClientID -SecurePassword $SecureEUMSecret
			        }
                    $Global:EUMSecret = (New-Object PSCredential "user",$SecureEUMSecret).GetNetworkCredential().Password
		        }
		        else {
			        $Global:EUMClientID = $EUMCredentials.UserName
			        $Global:EUMSecret = (New-Object PSCredential "user",$EUMCredentials.Password).GetNetworkCredential().Password
                    Write-Output "Connecting to EUM with Client ID" $EUMClientID
		        }
            }
	    }
	    else
	    {
		    Write-Output "Required Windows Credential Manager 2.0 PowerShell Module not found. Please install the module by entering the following command in PowerShell: ""Install-Module -Name ""CredentialManager"" -RequiredVersion 2.0"""
		    break
	    }
    }
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
            Write-Output "Exception Type: $($_.Exception.GetType().FullName)"
            Write-Output "Exception Message: $($_.Exception.Message)"
        }
    }
    catch
    {
        Write-Output "Exception Type: $($_.Exception.GetType().FullName)"
        Write-Output "Exception Message: $($_.Exception.Message)"
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
        [parameter(Mandatory=$false)]$parentURL,
        [parameter(Mandatory=$false)][string]$breadcrumbHTML,
        [parameter(Mandatory=$false)][string]$brandingDeploymentType,
        [parameter(Mandatory=$false)]$selectedTheme,
        [parameter(Mandatory=$false)]$siteTemplateName,
        [parameter(Mandatory=$false)]$siteCreatedDate,
        [parameter(Mandatory=$false)]$spSubWebs
    )

    Write-Output "Adding $($siteTitle) to the $($SiteListName) list. Please wait..."

    [hashtable]$newListItemValues = PrepareSiteItemValues -siteRelativeURL $siteRelativeURL -siteTitle $siteTitle -parentURL $parentURL `
        -breadcrumbHTML $breadcrumbHTML -brandingDeploymentType $brandingDeploymentType -selectedThemeName $selectedTheme.name `
        -masterPageName $selectedTheme.masterPage -siteTemplateName $siteTemplateName -siteCreatedDate $siteCreatedDate


    $existingItem = GetSiteEntry -siteRelativeURL $siteRelativeURL -disconnect

    Helper-Connect-PnPOnline -Url $SitesListSiteURL
    if ($existingItem)
    {
        Write-Output "$($siteTitle) exists in $($SiteListName) list. Updating..."
        [Microsoft.SharePoint.Client.ListItem]$newListItem = Set-PnPListItem -Identity $existingItem.Id -List $SiteListName -Values $newListItemValues -ContentType "EUM Site Collection List"
    }
    else
    {
        [Microsoft.SharePoint.Client.ListItem]$newListItem = Add-PnPListItem -List $SiteListName -Values $newListItemValues -ContentType "EUM Site Collection List"
    }

    if ($newListItem)
    {
        Write-Output "The site $($siteTitle) was added to the $($SiteListName) list successfully"
    }

    # -----------
    # Subsites
    # -----------
    if ($spSubWebs)
    {
        Write-Output "Adding subsites of $($siteTitle) to $($SiteListName). Please wait..."
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

            Helper-Connect-PnPOnline -Url $SitesListSiteURL
            if ($existingItem)
            {
                Write-Output "$($siteTitle) exists in $($SiteListName) list. Updating..."
                [Microsoft.SharePoint.Client.ListItem]$newListItem = Set-PnPListItem -Identity $existingItem.Id -List $SiteListName -Values $newListItemValues -ContentType "EUM Site Collection List"
            }
            else
            {
                [Microsoft.SharePoint.Client.ListItem]$newListItem = Add-PnPListItem -List $SiteListName -Values $newListItemValues -ContentType "EUM Site Collection List"
            }

            if ($newListItem)
            {
                Write-Output "The site $($siteTitle) was added to the $($SiteListName) list successfully"
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
        Helper-Connect-PnPOnline -Url $SitesListSiteURL

        Write-Output "Adding $($siteTitle) to the $($SiteListName) list. Please wait..."

        [hashtable]$newListItemValues = PrepareSiteItemValues -siteRelativeURL $siteRelativeURL -siteTitle $siteTitle -parentURL $parentURL `
            -breadcrumbHTML $breadcrumbHTML -brandingDeploymentType $brandingDeploymentType -selectedThemeName $selectedTheme.name `
            -masterPageName $selectedTheme.masterPage -siteTemplateName $siteTemplateName -siteCreatedDate $siteCreatedDate -subSite $false
        [Microsoft.SharePoint.Client.ListItem]$newListItem = Add-PnPListItem -List $SiteListName -Values $newListItemValues -ContentType "EUM Site Collection List"
    

        if ($newListItem)
        {
            Write-Output "The site $($siteTitle) was added to the $($SiteListName) list successfully"
        }

        # -----------
        # Subsites
        # -----------
        if ($spSubWebs)
        {
            Write-Output "Adding subsites of $($siteTitle) to $($SiteListName). Please wait..."
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

                Helper-Connect-PnPOnline -Url $SitesListSiteURL
                if ($existingItem)
                {
                    Write-Output "$($siteTitle) exists in $($SiteListName) list. Updating..."
                    [Microsoft.SharePoint.Client.ListItem]$newListItem = Set-PnPListItem -Identity $existingItem.Id -List $SiteListName -Values $newListItemValues -ContentType "EUM Site Collection List"
                }
                else
                {
                    [Microsoft.SharePoint.Client.ListItem]$newListItem = Add-PnPListItem -List $SiteListName -Values $newListItemValues -ContentType "EUM Site Collection List"
                }

                if ($newListItem)
                {
                    Write-Output "The site $($siteTitle) was added to the $($SiteListName) list successfully"
                }
            }
        }
    }
	else
	{
		Write-Output "The site $($siteTitle) exists in $($SiteListName) list. Skipping..."
	}
}