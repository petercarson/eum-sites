function LoadEnvironmentSettings() {

    [string]$Global:pnpTemplatePath = "c:\pnptemplates"

    # Check if running in Azure Automation or locally
    $Global:AzureAutomation = (Get-Command "Get-AutomationVariable" -errorAction SilentlyContinue)
    if ($AzureAutomation) {
        # Get automation variables
        $Global:SPCredentials = Get-AutomationPSCredential -Name 'SPOnlineCredentials'

        [string]$Global:SiteListName = Get-AutomationVariable -Name 'SiteListName'
        [string]$Global:TeamsChannelsListName = Get-AutomationVariable -Name 'TeamsChannelsListName'
        [string]$Global:WebAppURL = Get-AutomationVariable -Name 'WebAppURL'
        [string]$Global:AdminURL = $WebAppURL.Replace(".sharepoint.com", "-admin.sharepoint.com")
        [string]$Global:SitesListSiteURL = "$($WebAppURL)$(Get-AutomationVariable -Name 'SitesListSiteURL')"
        [string]$Global:SiteCollectionAdministrator = Get-AutomationVariable -Name 'siteCollectionAdministrator'
        [string]$Global:TeamsSPFxAppId = Get-AutomationVariable -Name 'TeamsSPFxAppId'


        $Global:AADCredentials = (Get-AutomationPSCredential -Name 'AADCredentials' -ErrorAction SilentlyContinue)
        if ($AADCredentials -ne $null) {
            $Global:AADClientID = $AADCredentials.UserName
            $Global:AADSecret = (New-Object PSCredential "user", $AADCredentials.Password).GetNetworkCredential().Password
            $Global:AADDomain = Get-AutomationVariable -Name 'AADDomain'
        }
    }
    else {
        [xml]$config = Get-Content -Path "$DistributionFolder\sharepoint.config"

        [System.Array]$Global:managedPaths = $config.settings.common.managedPaths.path
        [string]$Global:SiteListName = $config.settings.common.siteLists.siteListName
        [string]$Global:TeamsChannelsListName = $config.settings.common.siteLists.teamsChannelsListName

        $environmentId = $config.settings.common.defaultEnvironment

        if (-not $environmentId) {
            # Get the value from the last run as a default
            if ($environment.id) {
                $defaultText = "(Default - $($environment.id))"
            }

            #-----------------------------------------------------------------------
            # Prompt for the environment defined in the config
            #-----------------------------------------------------------------------

            Write-Verbose -Verbose -Message "`n***** AVAILABLE ENVIRONMENTS *****"
            $config.settings.environments.environment | ForEach {
                Write-Verbose -Verbose -Message "$($_.id)`t $($_.name) - $($_.webApp.URL)"
            }
            Write-Verbose -Verbose -Message "***** AVAILABLE ENVIRONMENTS *****"

            Do {
                [int]$environmentId = Read-Host "Enter the ID of the environment from the above list $defaultText"
            }
            Until (($environmentId -gt 0) -or ($environment.id -gt 0))
        }

        if ($environmentId -eq 0) {
            $environmentId = $environment.id
        }

        [System.Xml.XmlLinkedNode]$Global:environment = $config.settings.environments.environment | Where { $_.id -eq $environmentId }

        # Set variables based on environment selected
        [string]$Global:WebAppURL = $environment.webApp.url
        [string]$Global:AdminURL = $environment.webApp.url.Replace(".sharepoint.com", "-admin.sharepoint.com")
        [string]$Global:SitesListSiteURL = "$($WebAppURL)$($environment.webApp.sitesListSiteCollectionPath)"
        [string]$Global:SiteCollectionAdministrator = $environment.webApp.siteCollectionAdministrator
        [string]$Global:TeamsSPFxAppId = $environment.webApp.teamsSPFxAppId
        
        Write-Verbose -Verbose -Message "Environment set to $($environment.name) - $($environment.webApp.URL) `n"

        $Global:SPCredentials = GetManagedCredentials -managedCredentials $environment.webApp.managedCredentials -ManagedCredentialsType $environment.webApp.managedCredentialsType

        $AADCredentials = GetManagedCredentials -managedCredentials $environment.graphAPI.managedCredentials -ManagedCredentialsType $environment.graphAPI.managedCredentialsType
        if ($AADCredentials -ne $null) {
            $Global:AADClientID = $AADCredentials.UserName
            $Global:AADSecret = (New-Object PSCredential "user", $AADCredentials.Password).GetNetworkCredential().Password
            $Global:AADDomain = $environment.graphAPI.AADDomain
        }
    }
}

function GetManagedCredentials() {
    [OutputType([System.Management.Automation.PSCredential])]
    Param
    (
        [Parameter(Mandatory = $true)][string] $managedCredentials,
        [Parameter(Mandatory = $true)][string] $managedCredentialsType
    )

    if (-not(Get-InstalledModule -Name "CredentialManager" -RequiredVersion "2.0")) {
        Write-Verbose -Verbose -Message "Required Windows Credential Manager 2.0 PowerShell Module not found. Please install the module by entering the following command in PowerShell: ""Install-Module -Name ""CredentialManager"" -RequiredVersion 2.0"""
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
                Write-Verbose -Verbose -Message "Connecting with username $($Credentials.UserName)" 
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
                Write-Verbose -Verbose -Message "Connecting with Client Id $($Credentials.UserName)" 
            }
        }
    }

    return ($Credentials)
}

function Helper-Connect-PnPOnline() {
    Param
    (
        [Parameter(Mandatory = $true)][string] $URL
    )

    if ($O365ClientID -and $O365ClientSecret) {
        Connect-PnPOnline -Url $URL -AppId $O365ClientID -AppSecret $O365ClientSecret
    }
    else {
        Connect-PnPOnline -Url $URL -Credentials $SPCredentials
    }
}

function GetBreadcrumbHTML() {
    Param
    (
        [Parameter(Mandatory = $true)][string] $siteURL,
        [Parameter(Mandatory = $true)][string] $siteTitle,
        [Parameter(Mandatory = $false)][string] $parentURL
    )
    [string]$parentBreadcrumbHTML = ""

    if ($parentURL) {
        Helper-Connect-PnPOnline -Url $SitesListSiteURL

        $parentListItem = Get-PnPListItem -List $SiteListName -Query "
				<View>
						<Query>
								<Where>
										<Eq>
												<FieldRef Name='EUMSiteURL'/>
												<Value Type='Text'>$($parentURL)</Value>
										</Eq>
								</Where>
						</Query>
				</View>"

        if ($parentListItem) {
            [string]$parentBreadcrumbHTML = $parentListItem["EUMBreadcrumbHTML"]
        }
        else {
            Write-Verbose -Verbose -Message "No entry found for $parentURL"
        }
    }

    $siteURL = $siteURL.Replace($webAppURL, "")
    [string]$breadcrumbHTML = "<a href=`"$($siteURL)`">$($siteTitle)</a>"
    if ($parentBreadcrumbHTML) {
        $breadcrumbHTML = $parentBreadcrumbHTML + ' &gt; ' + $breadcrumbHTML
    }
    return $breadcrumbHTML
}

function PrepareSiteItemValues() {
    Param
    (
        [parameter(Mandatory = $false)][string]$siteURL,
        [parameter(Mandatory = $false)][string]$siteTitle,
        [parameter(Mandatory = $false)]$parentURL,
        [parameter(Mandatory = $false)][string]$breadcrumbHTML,
        [parameter(Mandatory = $false)]$siteCreatedDate,
        [parameter(Mandatory = $false)]$subSite
    )

    [hashtable]$newListItemValues = @{ }

    if ($siteURL) {
        $newListItemValues.Add("EUMSiteURL", $siteURL)
    }

    if ($siteTitle) {
        $newListItemValues.Add("Title", $siteTitle)
    }

    if ($parentURL) {
        $newListItemValues.Add("EUMParentURL", $parentURL)
    }

    if ($breadcrumbHTML) {
        $newListItemValues.Add("EUMBreadcrumbHTML", $breadcrumbHTML)
    }

    if ($siteCreatedDate) {
        $newListItemValues.Add("EUMSiteCreated", $siteCreatedDate)
    }

    if ($subSite) {
        $newListItemValues.Add("EUMSubSite", $subSite)
    }

    return $newListItemValues
}

function GetGraphAPIBearerToken() {
    $scope = "https://graph.microsoft.com/.default"
    $authorizationUrl = "https://login.microsoftonline.com/$($AADDomain)/oauth2/v2.0/token"

    Add-Type -AssemblyName System.Web

    $requestBody = @{
        client_id     = $AADClientID
        client_secret = $AADSecret
        scope         = $scope
        grant_type    = 'client_credentials'
    }

    $request = @{
        ContentType = 'application/x-www-form-urlencoded'
        Method      = 'POST'
        Body        = $requestBody
        Uri         = $authorizationUrl
    }

    $response = Invoke-RestMethod @request

    return $response.access_token
}

function GetGraphAPIServiceAccountBearerToken() {
    $scope = "https://graph.microsoft.com/.default"
    $authorizationUrl = "https://login.microsoftonline.com/$($AADDomain)/oauth2/v2.0/token"

    Add-Type -AssemblyName System.Web

    $requestBody = @{
        client_id     = $AADClientID
        client_secret = $AADSecret
        scope         = $scope
        grant_type    = 'password'
        username      = "$($SPCredentials.UserName)"
        password      = "$($SPCredentials.GetNetworkCredential().Password)"
    }

    $request = @{
        ContentType = 'application/x-www-form-urlencoded'
        Method      = 'POST'
        Body        = $requestBody
        Uri         = $authorizationUrl
    }

    $response = Invoke-RestMethod @request

    return $response.access_token
}

function AddOneNoteTeamsChannelTab() {
    Param
    (
        [parameter(Mandatory = $true)]$groupId,
        [parameter(Mandatory = $true)]$channelName,
        [parameter(Mandatory = $true)]$teamsChannelId,
        [parameter(Mandatory = $true)]$siteURL
    )

    $graphApiBaseUrl = "https://graph.microsoft.com/v1.0"

    # Retrieve access token for graph API
    $accessToken = GetGraphAPIBearerToken

    # Call the Graph API to get the notebook
    Write-Verbose -Verbose -Message "Retrieving notebook for group $($groupId)..."
    $graphGETEndpoint = "$($graphApiBaseUrl)/groups/$($groupId)/onenote/notebooks"

    # The notebook is not immediately available when the team site is created so use retry logic
    $getResponse = $null 
    while (($retries -lt 120) -and ($getResponse -eq $null -or $getResponse.value -eq $null)) {
        Start-Sleep -Seconds 30
        $retries += 1
        $getResponse = Invoke-RestMethod -Headers @{Authorization = "Bearer $accessToken" } -Uri $graphGETEndpoint -Method Get -ContentType 'application/json'
    }

    if ($getResponse -ne $null -and $getResponse.value -ne $null) {
        $notebookId = $getResponse.value.id
        $oneNoteWebUrl = $getResponse.value.links.oneNoteWebUrl

        # Call the Graph API to create a OneNote section
        Write-Verbose -Verbose -Message "Adding section $($channelName) to notebook for group $($groupId)..."
        $graphPOSTEndpoint = "$($graphApiBaseUrl)/groups/$($groupId)/onenote/notebooks/$($notebookId)/sections"
        $graphPOSTBody = @{
            "displayName" = $channelName
        }
        $postResponse = Invoke-RestMethod -Headers @{Authorization = "Bearer $accessToken" } -Uri $graphPOSTEndpoint -Body $($graphPOSTBody | ConvertTo-Json) -Method Post -ContentType 'application/json'
        $sectionId = $postResponse.id

        # Add a blank page to the section created above (required in order to link to the section)
        Write-Verbose -Verbose -Message "Adding page to section $($channelName) in notebook..."
        $graphPOSTEndpoint = "$($graphApiBaseUrl)/groups/$($groupId)/onenote/sections/$($sectionId)/pages"
        $graphPOSTBody = "<!DOCTYPE html><html><head><title></title><meta name='created' content='" + $(Get-Date -Format s) + "' /></head><body></body></html>"
        $postResponse = Invoke-RestMethod -Headers @{Authorization = "Bearer $accessToken" } -Uri $graphPOSTEndpoint -Body $graphPOSTBody -Method Post -ContentType 'text/html'

        # Add a tab to the team channel to the OneNote section    
        Write-Verbose -Verbose -Message "Adding OneNote tab to channel $($channelName)..."
        $configurationProperties = @{
            "contentUrl" = "https://www.onenote.com/teams/TabContent?notebookSource=PickSection&notebookSelfUrl=https://www.onenote.com/api/v1.0/myOrganization/groups/$($groupId)/notes/notebooks/$($notebookId)&oneNoteWebUrl=$($oneNoteWebUrl)&notebookName=OneNote&siteUrl=$($siteURL)&createdTeamType=Standard&wd=target(//$($channelName).one|/)&sectionId=$($notebookId)9&notebookIsDefault=true&ui={locale}&tenantId={tid}"
            "removeUrl"  = "https://www.onenote.com/teams/TabRemove?notebookSource=PickSection&notebookSelfUrl=https://www.onenote.com/api/v1.0/myOrganization/groups/$($groupId)/notes/notebooks/$($notebookId)c&oneNoteWebUrl=$($oneNoteWebUrl)&notebookName=OneNote&siteUrl=$($siteURL)&createdTeamType=Standard&wd=target(//$($channelName).one|/)&sectionId=$($notebookId)9&notebookIsDefault=true&ui={locale}&tenantId={tid}"
            "websiteUrl" = "https://www.onenote.com/teams/TabRedirect?redirectUrl=$($oneNoteWebUrl)?wd=target(%2F%2F$($channelName).one%7C%2F)"
        }
        $graphPOSTBody = @{
            "teamsApp@odata.bind" = "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/0d820ecd-def2-4297-adad-78056cde7c78"
            "displayName"         = "OneNote"
            "configuration"       = $configurationProperties
        }
        $graphPOSTEndpoint = "$($graphApiBaseUrl)/teams/$($groupId)/channels/$($teamsChannelId)/tabs"
        $postResponse = Invoke-RestMethod -Headers @{Authorization = "Bearer $accessToken" } -Uri $graphPOSTEndpoint -Body $($graphPOSTBody | ConvertTo-Json) -Method Post -ContentType 'application/json'
    }
    else {
        Write-Error "Could not retrieve notebook for group $($groupId)"
    }
}


function AddTeamsChannelRequestFormToChannel() {
    Param
    (
        [parameter(Mandatory = $true)]$groupId,
        [parameter(Mandatory = $true)]$teamsChannelId
    )
    
    $graphApiBaseUrl = "https://graph.microsoft.com/v1.0"

    # Retrieve access token for graph API
    $accessToken = GetGraphAPIBearerToken

    # First add the app to the team
    Write-Verbose -Verbose -Message "Adding Add channel SPFx Web Part app to team for groupId $($groupId)..."
    $graphPOSTEndpoint = "$($graphApiBaseUrl)/teams/$($groupId)/installedApps"
    $graphPOSTBody = @{
        "teamsApp@odata.bind" = "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/$($TeamsSPFxAppId)"
        "id"                  = "$($TeamsSPFxAppId)"
        "externalId"          = "75dbe34f-74a5-4bbb-9495-41701c0d7ac0"
        "name"                = "Add channel"
        "version"             = "0.1"
        "distributionMethod"  = "organization"
    }
    $postResponse = Invoke-RestMethod -Headers @{Authorization = "Bearer $accessToken" } -Uri $graphPOSTEndpoint -Body $($graphPOSTBody | ConvertTo-Json) -Method Post -ContentType 'application/json'

    Start-Sleep -Seconds 60

    # Add the SPFx web part to the channel
    Write-Verbose -Verbose -Message "Adding Add channel SPFx Web Part tab to channel $($teamsChannelId)..."
    $graphPOSTEndpoint = "$($graphApiBaseUrl)/teams/$($groupId)/channels/$($teamsChannelId)/tabs"
    $graphPOSTBody = @{
        "displayName"         = "Add channel"
        "teamsApp@odata.bind" = "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/$($TeamsSPFxAppId)"
    }
    $postResponse = Invoke-RestMethod -Headers @{Authorization = "Bearer $accessToken" } -Uri $graphPOSTEndpoint -Body $($graphPOSTBody | ConvertTo-Json) -Method Post -ContentType 'application/json'
}

function AddGroupOwner() {
    Param
    (
        [parameter(Mandatory = $true)]$groupId,
        [parameter(Mandatory = $true)]$email
    )
    
    $graphApiBaseUrl = "https://graph.microsoft.com/v1.0"

    # Retrieve access token for graph API
    $accessToken = GetGraphAPIBearerToken

    Write-Verbose -Verbose -Message "Adding $($email) as owner to groupId $($groupId)..."
    $graphPOSTEndpoint = "$($graphApiBaseUrl)/groups/$($groupId)/owners/`$ref"
    $graphPOSTBody = @{
        "@odata.id" = "$($graphApiBaseUrl)/users/$($email)"
    }
    $postResponse = Invoke-RestMethod -Headers @{Authorization = "Bearer $accessToken" } -Uri $graphPOSTEndpoint -Body $($graphPOSTBody | ConvertTo-Json) -Method Post -ContentType 'application/json'
}

function AddTeamPlanner() {
    Param
    (
        [parameter(Mandatory = $true)]$groupId,
        [parameter(Mandatory = $true)]$planTitle
    )
    
    $graphApiBaseUrl = "https://graph.microsoft.com/v1.0"

    # Retrieve access token for graph API
    $accessToken = GetGraphAPIServiceAccountBearerToken
    Write-Verbose -Verbose -Message $accessToken

    Write-Verbose -Verbose -Message "Creating plan $($planTitle) for groupId $($groupId)..."
    $graphPOSTEndpoint = "$($graphApiBaseUrl)/planner/plans"
    $graphPOSTBody = @{
        "owner" = $($groupId)
        "title" = $($planTitle)
    }
    $postResponse = Invoke-RestMethod -Headers @{Authorization = "Bearer $accessToken" } -Uri $graphPOSTEndpoint -Body $($graphPOSTBody | ConvertTo-Json) -Method Post -ContentType 'application/json'

    return $postResponse.id
}

function AddPlannerTeamsChannelTab() {
    Param
    (
        [parameter(Mandatory = $true)]$groupId,
        [parameter(Mandatory = $true)]$planTitle,
        [parameter(Mandatory = $true)]$planId,
        [parameter(Mandatory = $true)]$channelName,
        [parameter(Mandatory = $true)]$teamsChannelId
    )

    $graphApiBaseUrl = "https://graph.microsoft.com/v1.0"

    # Retrieve access token for graph API
    $accessToken = GetGraphAPIBearerToken
    Write-Verbose -Verbose -Message $accessToken

    Write-Verbose -Verbose -Message "Adding Planner tab for plan $($planTitle) to channel $($channelName)..."
    $configurationProperties = @{
        "entityId"   = $planId
        "contentUrl" = "https://tasks.office.com/$($AADDomain)/Home/PlannerFrame?page=7&planId=$($planId)"
        "removeUrl"  = "https://tasks.office.com/$($AADDomain)/Home/PlannerFrame?page=7&planId=$($planId)"
        "websiteUrl" = "https://tasks.office.com/$($AADDomain)/Home/PlannerFrame?page=7&planId=$($planId)"
    }

    $graphPOSTBody = @{
        "teamsApp@odata.bind" = "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/com.microsoft.teamspace.tab.planner"
        "displayName"         = "$($planTitle) Planner"
        "configuration"       = $configurationProperties
    }

    $graphPOSTEndpoint = "$($graphApiBaseUrl)/teams/$($groupId)/channels/$($teamsChannelId)/tabs"
    $postResponse = Invoke-RestMethod -Headers @{Authorization = "Bearer $accessToken" } -Uri $graphPOSTEndpoint -Body $($graphPOSTBody | ConvertTo-Json) -Method Post -ContentType 'application/json'
}

Set-PnPTraceLog -On -Level Debug