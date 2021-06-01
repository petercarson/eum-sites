Param
(
    [Parameter (Mandatory = $false)][int]$testItemID = -1
)

import-module Az.Keyvault
#import-module Az.Automation

function LoadEnvironmentSettings() {

    [string]$Global:pnpTemplatePath = "c:\pnptemplates"
    [string]$Global:SiteListName = "Sites"
    [string]$Global:TeamsChannelsListName = "TeamsChannels"
    [string]$Global:TenantID = ""


    # Check if running in Azure Automation or locally
    $Global:AzureAutomation = (Get-Command "Get-AutomationVariable" -errorAction SilentlyContinue)
    if ($AzureAutomation) {
        [string]$Global:RootURL = Get-AutomationVariable -Name 'RootURL'
        [string]$Global:AdminURL = $RootURL.Replace(".sharepoint.com", "-admin.sharepoint.com")
        [string]$Global:SiteCollectionRelativeURL = Get-AutomationVariable -Name 'SiteCollectionRelativeURL'
        [string]$Global:SiteCollectionFullURL = "$($RootURL)$($SiteCollectionRelativeURL)"
        [string]$Global:SiteCollectionAdministrator = Get-AutomationVariable -Name 'siteCollectionAdministrator'
        [string]$Global:TeamsSPFxAppId = Get-AutomationVariable -Name 'TeamsSPFxAppId'
        #[string]$Global:AutomationAccountName = Get-AutomationVariable -Name 'AutomationAccountName'
        #[string]$Global:ResourceGroupName = Get-AutomationVariable -Name 'ResourceGroupName'
        [string]$Global:KeyVaultName = Get-AutomationVariable -Name 'KeyVaultName'

        [boolean]$Global:IsSharePointOnline = $RootURL.ToLower() -like "*.sharepoint.com"
    }
    else {
        [xml]$config = Get-Content -Path "$($PSScriptRoot)\sharepoint.config"

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

        if ($environmentId -eq 0) {
            $environmentId = $environment.id
        }

        [System.Xml.XmlLinkedNode]$Global:environment = $config.settings.environments.environment | Where { $_.id -eq $environmentId }

        # Set variables based on environment selected
        [string]$Global:RootURL = $environment.rootURL
        [string]$Global:AdminURL = $environment.rootURL.Replace(".sharepoint.com", "-admin.sharepoint.com")
        [string]$Global:SiteCollectionRelativeURL = $environment.siteCollectionRelativeURL
        [string]$Global:SiteCollectionFullURL = "$($RootURL)$($SiteCollectionRelativeURL)"
        [string]$Global:SiteCollectionAdministrator = $environment.siteCollectionAdministrator
        [string]$Global:TeamsSPFxAppId = $environment.teamsSPFxAppId
        [string]$Global:PrimaryDomain = $environment.PrimaryDomain
        [boolean]$Global:IsSharePointOnline = $RootURL.ToLower() -like "*.sharepoint.com"
        [string]$Global:CredentialsType = $environment.credentialsType

        Write-Verbose -Verbose -Message "Environment set to $($environment.name) - $($environment.webApp.URL) `n"

        switch ($CredentialsType) {
            "UsernamePassword" {
                if ($SPCredentials -eq $null) {
                    $Global:SPCredentials = Get-Credential
                }
            }
            "Certificate" {
                $Global:clientID = $environment.clientID
                $Global:thumbprint = $environment.thumbprint
            }
            "Interactive" {
            }
        }
    }
}

function GetGraphAccessTokenFromRefreshToken {
     param ([Parameter(Mandatory=$true)][string]$refreshToken,
            [Parameter(Mandatory=$true)][string]$clientId,
            [Parameter(Mandatory=$true)][string]$clientSecret
     )

	Try 
    {
        $redirectUrl = "https://localhost:8000"
        $resourceUrl = "https://graph.microsoft.com"

        # Add System Web Assembly to encode ClientSecret etc.
        Add-Type -AssemblyName System.Web

        # UrlEncode the ClientID and ClientSecret and URL's for special characters 
        $clientIDEncoded = [System.Web.HttpUtility]::UrlEncode($clientId)
        $clientSecretEncoded = [System.Web.HttpUtility]::UrlEncode($clientSecret)
        $redirectUrlEncoded =  [System.Web.HttpUtility]::UrlEncode($redirectUrl)
        $resourceUrlEncoded = [System.Web.HttpUtility]::UrlEncode($resourceUrl)

		$refreshBody = "grant_type=refresh_token&redirect_uri=$redirectUrlEncoded&client_id=$clientIdEncoded&client_secret=$clientSecretEncoded&refresh_token=$refreshToken&resource=$resourceUrlEncoded"

		$Authorization = Invoke-RestMethod https://login.microsoftonline.com/common/oauth2/token `
			-Method Post -ContentType "application/x-www-form-urlencoded" `
			-Body $refreshBody `
			-UseBasicParsing

        return $Authorization.access_token
	}
	Catch 
    {
        Write-Host "Error getting access token: '$($_.Exception.Message)'"
        return $null
	}
}

function Helper-Connect-PnPOnline() {
    Param
    (
        [Parameter(Mandatory = $true)][string] $URL
    )

    $Global:AzureAutomation = (Get-Command "Get-AutomationVariable" -errorAction SilentlyContinue)
    if ($AzureAutomation) {
        # Get Azure Run As Connection Name
        $connectionName = "AzureRunAsConnection"
        # Get the Service Principal connection details for the Connection name
        $servicePrincipalConnection = Get-AutomationConnection -Name $connectionName         

        $Conn = Connect-PnPOnline -Tenant $servicePrincipalConnection.TenantId -ClientId $servicePrincipalConnection.ApplicationId -Thumbprint $servicePrincipalConnection.CertificateThumbprint -Url $URL -ReturnConnection
        }
    else {
        switch ($CredentialsType) {
            "UsernamePassword" {
                $Conn = Connect-PnPOnline -Url $URL -Credentials $SPCredentials -ReturnConnection
            }
            "Certificate" {
                $Conn = Connect-PnPOnline -Tenant $PrimaryDomain -ClientId $clientID -Thumbprint $thumbprint -Url $URL -ReturnConnection
            }
            "Interactive" {
                $Conn = Connect-PnPOnline -Url $URL -Interactive -ReturnConnection
            }
        }
    }

    return $Conn
}

function Helper-Connect-AzAccount() {
    $Global:AzureAutomation = (Get-Command "Get-AutomationVariable" -errorAction SilentlyContinue)
    if ($AzureAutomation) {
        # Get Azure Run As Connection Name
        $connectionName = "AzureRunAsConnection"
        # Get the Service Principal connection details for the Connection name
        $servicePrincipalConnection = Get-AutomationConnection -Name $connectionName         
        Connect-AzAccount -CertificateThumbprint $servicePrincipalConnection.CertificateThumbprint -ApplicationId $servicePrincipalConnection.ApplicationId -TenantId $servicePrincipalConnection.TenantId
        $Global:TenantId = $servicePrincipalConnection.TenantId
    }
    else {
        switch ($CredentialsType) {
            "UsernamePassword" {
                $AzureADConnection = Connect-AzAccount -Credential $SPCredentials
            }
            "Certificate" {
                $AzureADConnection = Connect-AzAccount -CertificateThumbprint $thumbprint -ApplicationId $clientID -TenantId $PrimaryDomain
            }
            "Interactive" {
                $AzureADConnection = Connect-AzAccount
            }
        }
        $Global:TenantID = $AzureADConnection.TenantId
    }
}

function Helper-Connect-MicrosoftTeams() {
    $Global:AzureAutomation = (Get-Command "Get-AutomationVariable" -errorAction SilentlyContinue)
    if ($AzureAutomation) {
        # Get Azure Run As Connection Name
        $connectionName = "AzureRunAsConnection"
        # Get the Service Principal connection details for the Connection name
        $servicePrincipalConnection = Get-AutomationConnection -Name $connectionName         

        Connect-MicrosoftTeams -TenantId $servicePrincipalConnection.TenantId -ApplicationId $servicePrincipalConnection.ApplicationId -CertificateThumbprint $servicePrincipalConnection.CertificateThumbprint
        }
    else {
        switch ($CredentialsType) {
            "UsernamePassword" {
                Connect-MicrosoftTeams -TenantId $PrimaryDomain -Credential $SPCredentials
            }
            "Certificate" {
                Connect-MicrosoftTeams -TenantId $PrimaryDomain -ApplicationId $clientID -CertificateThumbprint $thumbprint
            }
            "Interactive" {
                Connect-MicrosoftTeams
            }
        }
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
        $connLandingSite = Helper-Connect-PnPOnline -Url $SiteCollectionFullURL

        $parentListItem = Get-PnPListItem -List $SiteListName -Connection $connLandingSite -Query "
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

    $siteURL = $siteURL.Replace($RootURL, "")
    [string]$breadcrumbHTML = "<a href=`"$($siteURL)`">$($siteTitle)</a>"
    if ($parentBreadcrumbHTML) {
        $breadcrumbHTML = $parentBreadcrumbHTML + ' &gt; ' + $breadcrumbHTML
    }
    return $breadcrumbHTML
}

function AddOneNoteTeamsChannelTab() {
    Param
    (
        [parameter(Mandatory = $true)]$groupId,
        [parameter(Mandatory = $true)]$teamName,
        [parameter(Mandatory = $true)]$channelName,
        [parameter(Mandatory = $true)]$teamsChannelId,
        [parameter(Mandatory = $true)]$siteURL
    )

    Helper-Connect-AzAccount
    $graphApiBaseUrl = "https://graph.microsoft.com/v1.0"

    # Retrieve access token for graph API
    $accessToken = Get-PnPGraphAccessToken

    # Call the Graph API to get the notebook
    Write-Verbose -Verbose -Message "Retrieving notebook for group $($groupId)..."
    $graphGETEndpoint = "$($graphApiBaseUrl)/groups/$($groupId)/onenote/notebooks"

    # The notebook is not immediately available when the team site is created check if it exists and create it if necessary
    $getResponse = Invoke-RestMethod -Headers @{Authorization = "Bearer $accessToken" } -Uri $graphGETEndpoint -Method Get -ContentType 'application/json'

    if ($getResponse -eq $null -or $getResponse.value -eq $null -or $getResponse.value.length -eq 0) {
        $graphPOSTEndpoint = "$($graphApiBaseUrl)/groups/$($groupId)/onenote/notebooks"
        $graphPOSTBody = @{
            "displayName" = "$($teamName) Notebook"
        }
        Write-Verbose -Verbose -Message "Creating notebook for group $($groupId)..."
        $postResponse = Invoke-RestMethod -Headers @{Authorization = "Bearer $accessToken" } -Uri $graphPOSTEndpoint -Body $($graphPOSTBody | ConvertTo-Json) -Method Post -ContentType 'application/json'

        $notebookId = $postResponse.id
        $oneNoteWebUrl = $getResponse.links.oneNoteWebUrl
    }
    else {
        $notebookId = $getResponse.value[0].id
        $oneNoteWebUrl = $getResponse.value[0].links.oneNoteWebUrl
    }

    if ($notebookId -ne $null) {
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
        Write-Error "Could not retrieve or create notebook for group $($groupId)"
    }
}

function AddTeamPlanner() {
    Param
    (
        [parameter(Mandatory = $true)]$groupId,
        [parameter(Mandatory = $true)]$planTitle
    )
    
    $graphApiBaseUrl = "https://graph.microsoft.com/v1.0"

    $junk = Helper-Connect-AzAccount

    # Retrieve user delegate access token for graph API access to Planner 
    $refreshToken = Get-AzKeyVaultSecret -VaultName $KeyVaultName -Name "PlannerRefreshToken" -AsPlainText
    $credential = Get-AutomationPSCredential -Name "PlannerClient"
    $accessToken = GetGraphAccessTokenFromRefreshToken -refreshToken $refreshToken -clientId $Credential.Username -clientSecret $Credential.GetNetworkCredential().Password

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
    $accessToken = Get-PnPGraphAccessToken

    #Write-Verbose -Verbose -Message $accessToken

    Write-Verbose -Verbose -Message "Adding Planner tab for plan $($planTitle) and planId $($planId) to channel $($channelName)..."
    $configurationProperties = @{
        "entityId"   = $planId
        "contentUrl" = "https://tasks.office.com/$($TenantID)/Home/PlannerFrame?page=7&planId=$($planId)"
        "removeUrl"  = "https://tasks.office.com/$($TenantID)/Home/PlannerFrame?page=7&planId=$($planId)"
        "websiteUrl" = "https://tasks.office.com/$($TenantID)/Home/PlannerFrame?page=7&planId=$($planId)"
    }

    $graphPOSTBody = @{
        "teamsApp@odata.bind" = "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/com.microsoft.teamspace.tab.planner"
        "displayName"         = "Planner"
        "configuration"       = $configurationProperties
    }

    $graphPOSTEndpoint = "$($graphApiBaseUrl)/teams/$($groupId)/channels/$($teamsChannelId)/tabs"
    $postResponse = Invoke-RestMethod -Headers @{Authorization = "Bearer $accessToken" } -Uri $graphPOSTEndpoint -Body $($graphPOSTBody | ConvertTo-Json) -Method Post -ContentType 'application/json'
}

function GetGroupIdByName() {
    Param
    (
        [parameter(Mandatory = $true)]$groupName
    )

    $graphApiBaseUrl = "https://graph.microsoft.com/v1.0"

    # Retrieve access token for graph API
    $accessToken = Get-PnPGraphAccessToken
    $groupFormatted = $groupName.replace("'", "''")
    Write-Verbose -Verbose -Message "Retrieving group ID for group $($groupFormatted)..."
    $graphGETEndpoint = "$($graphApiBaseUrl)/groups?`$filter=displayName eq '$($groupFormatted)'"

    try {
        $getResponse = Invoke-RestMethod -Headers @{Authorization = "Bearer $accessToken" } -Uri $graphGETEndpoint -Method Get -ContentType 'application/json'
        Write-Verbose -Verbose -Message "Retrieving group ID $($getResponse.value.id) for group $($groupName)."
        return $getResponse.value.id
    }
    catch [System.Net.WebException] {
        if ([int]$_.Exception.Response.StatusCode -eq 404) {
            return $null
        }
        else {
            Write-Error "Exception Type: $($_.Exception.GetType().FullName)"
            Write-Error "Exception Message: $($_.Exception.Message)"
        }
    }
    catch {
        Write-Error "Exception Type: $($_.Exception.GetType().FullName)"
        Write-Error "Exception Message: $($_.Exception.Message)"
    }
}

function GetGroupIdByAlias() {
    Param
    (
        [parameter(Mandatory = $true)]$groupAlias
    )

    $graphApiBaseUrl = "https://graph.microsoft.com/v1.0"

    # Retrieve access token for graph API
    $accessToken = Get-PnPGraphAccessToken
    Write-Verbose -Verbose -Message "Retrieving group ID for group $($groupAlias)..."
    $graphGETEndpoint = "$($graphApiBaseUrl)/groups?`$filter=mailNickname eq '$($groupAlias)'"

    try {
        $getResponse = Invoke-RestMethod -Headers @{Authorization = "Bearer $accessToken" } -Uri $graphGETEndpoint -Method Get -ContentType 'application/json'
        Write-Verbose -Verbose -Message "Retrieving group ID $($getResponse.value.id) for group $($groupAlias)."
        return $getResponse.value.id
    }
    catch [System.Net.WebException] {
        if ([int]$_.Exception.Response.StatusCode -eq 404) {
            return $null
        }
        else {
            Write-Error "Exception Type: $($_.Exception.GetType().FullName)"
            Write-Error "Exception Message: $($_.Exception.Message)"
        }
    }
    catch {
        Write-Error "Exception Type: $($_.Exception.GetType().FullName)"
        Write-Error "Exception Message: $($_.Exception.Message)"
    }
}

function ConvertGroupNameToAlias() {
    Param
    (
        [parameter(Mandatory = $true)]$groupName
    )
	[string]$groupAlias = $groupName.Replace(' ', '-')
    # https://docs.microsoft.com/en-us/office/troubleshoot/error-messages/username-contains-special-character
    # Convert any accented characters
    $groupAlias = [Text.Encoding]::ASCII.GetString([Text.Encoding]::GetEncoding("Cyrillic").GetBytes($groupAlias))
    # Remove any special characters
    $groupAlias = $groupAlias -replace '[^a-zA-Z0-9\-]', ''

    return $groupAlias
}

function GetGroupSiteUrl() {
    Param
    (
        [parameter(Mandatory = $true)]$groupId
    )

    $graphApiBaseUrl = "https://graph.microsoft.com/v1.0"

    # Retrieve access token for graph API
    $accessToken = Get-PnPGraphAccessToken

    Write-Verbose -Verbose -Message "Retrieving site URL for group $($groupId)..."
    $graphGETEndpoint = "$($graphApiBaseUrl)/groups/$($groupId)/sites/root/webUrl"

    try {
        $getResponse = Invoke-RestMethod -Headers @{Authorization = "Bearer $accessToken" } -Uri $graphGETEndpoint -Method Get -ContentType 'application/json'
        return $getResponse.value
    }
    catch [System.Net.WebException] {
        if ([int]$_.Exception.Response.StatusCode -eq 404) {
            return $null
        }
        else {
            Write-Error "Exception Type: $($_.Exception.GetType().FullName)"
            Write-Error "Exception Message: $($_.Exception.Message)"
        }
    }
    catch {
        Write-Error "Exception Type: $($_.Exception.GetType().FullName)"
        Write-Error "Exception Message: $($_.Exception.Message)"
    }
}

function ProvisionSite {
    Param
    (
        [Parameter (Mandatory = $True)][int]$listItemID
    )

    Write-Verbose -Verbose -Message "listItemID = $($listItemID)"

    $connLandingSite = Helper-Connect-PnPOnline -Url $SiteCollectionFullURL

    New-Item -Path $pnpTemplatePath -ItemType "directory" -Force | out-null

    $pnpTemplates = Find-PnPFile -List "PnPTemplates" -Match *.xml -Connection $connLandingSite
    $pnpTemplates | ForEach-Object {
        $File = Get-PnPFile -Url "$($SiteCollectionRelativeURL)/pnptemplates/$($_.Name)" -Path $pnpTemplatePath -Filename $_.Name -AsFile -Force -Connection $connLandingSite
    }

    # Get the specific Site Collection List item in master site for the site that needs to be created
    $pendingSite = Get-PnPListItem -Connection $connLandingSite -List $SiteListName -Query "
    <View>
        <Query>
            <Where>
                <And>
                    <And>
                        <Eq>
                            <FieldRef Name='ID'/>
                            <Value Type='Integer'>$listItemID</Value>
                        </Eq>
                        <IsNull>
                            <FieldRef Name='EUMSiteCreated'/>
                        </IsNull>
                    </And>
                    <Eq>
                        <FieldRef Name='_ModerationStatus' />
                        <Value Type='ModStat'>0</Value>
                    </Eq>
                </And>
            </Where>
        </Query>
        <ViewFields>
            <FieldRef Name='ID'></FieldRef>
            <FieldRef Name='Title'></FieldRef>
            <FieldRef Name='EUMSiteURL'></FieldRef>
            <FieldRef Name='EUMAlias'></FieldRef>
            <FieldRef Name='EUMSiteVisibility'></FieldRef>
            <FieldRef Name='EUMBreadcrumbHTML'></FieldRef>
            <FieldRef Name='EUMParentURL'></FieldRef>
            <FieldRef Name='EUMSiteTemplate'></FieldRef>
            <FieldRef Name='EUMDivision'></FieldRef>
            <FieldRef Name='EUMCreateTeam'></FieldRef>
            <FieldRef Name='EUMCreateOneNote'></FieldRef>
            <FieldRef Name='EUMCreatePlanner'></FieldRef>
            <FieldRef Name='EUMExternalSharing'></FieldRef>
            <FieldRef Name='EUMDefaultSharingLinkType'></FieldRef>
            <FieldRef Name='EUMDefaultLinkPermission'></FieldRef>
            <FieldRef Name='Requestor'></FieldRef>
        </ViewFields>
    </View>"

    if ($pendingSite.Count -gt 0) {
        # Get the time zone of the master site
        $spWeb = Get-PnPWeb -Includes RegionalSettings.TimeZone -Connection $connLandingSite
        [int]$timeZoneId = $spWeb.RegionalSettings.TimeZone.Id

        [string]$siteTitle = $pendingSite["Title"]
        [string]$alias = $pendingSite["EUMAlias"]
        if ($alias) {
            # Replace spaces in Alias with dashes
            $alias = $alias -replace '\s', '-'
            $siteURL = "$($RootURL)/sites/$alias"
        }
        else {
            [string]$siteURL = "$($RootURL)$($pendingSite['EUMSiteURL'])"
        }

        [string]$siteVisibility = $pendingSite["EUMSiteVisibility"]

        [boolean]$eumCreateTeam = $false
        if ($pendingSite["EUMCreateTeam"] -ne $null) { 
            $eumCreateTeam = $pendingSite["EUMCreateTeam"] 
        }

        [boolean]$eumCreateOneNote = $false 
        if ($pendingSite["EUMCreateOneNote"] -ne $null) {
            $eumCreateOneNote = $pendingSite["EUMCreateOneNote"]
        }

        [boolean]$eumCreatePlanner = $false 
        if ($pendingSite["EUMCreatePlanner"] -ne $null) {
            $eumCreatePlanner = $pendingSite["EUMCreatePlanner"]
        }

        [string]$breadcrumbHTML = $pendingSite["EUMBreadcrumbHTML"]
        [string]$parentURL = $pendingSite["EUMParentURL"]
        [string]$eumExternalSharing = $pendingSite["EUMExternalSharing"]
        [string]$eumDefaultSharingLinkType = $pendingSite["EUMDefaultSharingLinkType"]
        [string]$eumDefaultLinkPermission = $pendingSite["EUMDefaultLinkPermission"]
        [string]$Division = $pendingSite["EUMDivision"].LookupValue
        [string]$eumSiteTemplate = $pendingSite["EUMSiteTemplate"].LookupValue
        [string]$owner = $pendingSite["Requestor"].Email
        [string]$requestor = $owner

        [boolean]$parentHubSite = $false
        
        $divisionSiteURL = Get-PnPListItem -Connection $connLandingSite -List "Divisions" -Query "
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
																<FieldRef Name='HubSite'></FieldRef>
															</ViewFields>
														</View>"
		
        if ($divisionSiteURL.Count -eq 1) {
            if ($parentURL -eq "") { 
                $parentURL = $divisionSiteURL["SiteURL"].Url 
            }

            if (($divisionSiteURL["HubSite"] -ne "") -and ($divisionSiteURL["HubSite"] -ne $null)) {
                $parentHubSite = $divisionSiteURL["HubSite"]
            }
        }

        $siteTemplate = Get-PnPListItem -Connection $connLandingSite -List "Site Templates" -Query "
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
														<FieldRef Name='JoinHubSite'></FieldRef>
													</ViewFields>
												</View>"
		
        $baseSiteTemplate = ""
        $baseSiteType = ""
        $pnpSiteTemplate = ""
        $joinHubSite = $false
        $siteCreated = $false

        if ($siteTemplate.Count -eq 1) {
            $baseSiteTemplate = $siteTemplate["BaseClassicSiteTemplate"]
            $baseSiteType = $siteTemplate["BaseModernSiteType"]

            if ($siteTemplate["JoinHubSite"] -ne $null) { 
                $joinHubSite = $siteTemplate["JoinHubSite"] 
            }

            if ($siteTemplate["PnPSiteTemplate"] -ne $null) {
                $pnpSiteTemplate = "$pnpTemplatePath\$($siteTemplate["PnPSiteTemplate"].LookupValue)"
            }
        }

        # Classic style sites
        if ($baseSiteTemplate) {
            # Create the site
            if ($siteCollection) {
                # Create site (if it exists, it will error but not modify the existing site)
                Write-Verbose -Verbose -Message "Creating site collection $($siteURL) with base template $($baseSiteTemplate). Please wait..."
                try {
                    New-PnPTenantSite -Title $siteTitle -Url $siteURL -Owner $owner -TimeZone $timeZoneId -Template $baseSiteTemplate -RemoveDeletedSite -Wait -Force -Connection $connLandingSite -ErrorAction Stop
                }
                catch { 
                    Write-Error "Failed creating site collection $($siteURL)"
                    Write-Error $_
                }
            }
            else {
                # Connect to parent site
                $connParentSite = Helper-Connect-PnPOnline -Url $parentURL

                # Create the subsite
                Write-Verbose -Verbose -Message "Creating subsite $($siteURL) with base template $($baseSiteTemplate) under $($parentURL). Please wait..."

                [string]$subsiteURL = $siteURL.Replace($parentURL, "").Trim('/')
                New-PnPWeb -Title $siteTitle -Url $subsiteURL -Template $baseSiteTemplate -Connection $connParentSite

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

                        if ($IsSharePointOnline) {
                            $siteURL = New-PnPSite -Type CommunicationSite -Title $siteTitle -Url $siteURL -ErrorAction Stop -Connection $connLandingSite -Owner $owner
                        }
                        else {
                            New-PnPTenantSite -Title $siteTitle -Url $siteURL -Owner $owner -TimeZone $timeZoneId -Template "SITEPAGEPUBLISHING#0" -Wait -Force -Connection $connLandingSite -ErrorAction Stop
                        }
                        $siteCreated = $true
                    }
                    catch { 
                        Write-Error "Failed creating site collection $($siteURL)"
                        Write-Error $_
                        return $false
                    }
                }
                "TeamSite" {
                    try {
                        Write-Verbose -Verbose -Message "Creating site collection $($siteURL) with modern type $($baseSiteType). Please wait..."
                        if ($IsSharePointOnline) {
                            if ($eumCreateTeam) {
                                $teamVisibility = $siteVisibility
                                if ($teamVisibility -eq "Hidden") {
                                    $teamVisibility = "Private"
                                }
                                
                                Helper-Connect-MicrosoftTeams
                                $team = New-Team -DisplayName $siteTitle -MailNickName $alias -Owner $owner -Visibility $teamVisibility

                                $groupId = $team.GroupId
                                $teamsChannels = Get-TeamChannel -GroupId $groupId
                                $generalChannel = $teamsChannels | Where-Object { $_.DisplayName -eq 'General' }
                                $generalChannelId = $generalChannel.Id

                            }
                            else {
                                if ($siteVisibility -eq "Public") {
                                    $siteURL = New-PnPSite -Type TeamSite -Title $siteTitle -Alias $alias -IsPublic -Connection $connLandingSite -ErrorAction Stop -Owner $owner
                                }
                                else {
                                    $siteURL = New-PnPSite -Type TeamSite -Title $siteTitle -Alias $alias -Connection $connLandingSite -ErrorAction Stop -Owner $owner
                                }
                            }
                        }
                        else {
                            New-PnPTenantSite -Title $siteTitle -Url $siteURL -Owner $owner -TimeZone $timeZoneId -Template "STS#3" -Wait -Force -Connection $connLandingSite -ErrorAction Stop
                        }
                        $siteCreated = $true

                        
                    }
                    catch { 
                        Write-Error "Failed creating site collection $($siteURL)"
                        Write-Error $_
                        return $false
                    }
                }
            }
        }

        if ($siteCreated) {
            $connSite = Helper-Connect-PnPOnline -Url $siteURL

            #Set the external sharing capabilities 
            if ($eumExternalSharing){
                switch ($eumExternalSharing){
                    'Anyone' {$externalSharingOption = "ExternalUserAndGuestSharing"  ; Break}
                    'New and existing guests' {$externalSharingOption = "ExternalUserSharingOnly" ; Break}
                    'Existing guests only' {$externalSharingOption = "ExistingExternalUserSharingOnly" ; Break}
                    'Only people in your organization' {$externalSharingOption = "Disabled" ; Break}
                }

                Write-Verbose -Verbose -Message "Setting external sharing to $($externalSharingOption)"
                Set-PnPSite -Identity $siteURL -Sharing $externalSharingOption
            }

            #Set the default sharing link type 
            if ($eumDefaultSharingLinkType){
                 switch ($eumDefaultSharingLinkType){
                    'Anyone with the link' {$defaultSharingLinkTypeOption = "AnonymousAccess" ; Break}
                    'Specific people' {$defaultSharingLinkTypeOption = "Direct" ; Break}
                    'Only people in your organization' {$defaultSharingLinkTypeOption = "Internal "; Break}
                    'People with existing access' {$defaultSharingLinkTypeOption = "ExistingAccess"; Break}  
                }
                if ($defaultSharingLinkTypeOption -and $defaultSharingLinkTypeOption -ne "ExistingAccess"){
                    Set-PnPSite -Identity $siteURL -DefaultSharingLinkType $defaultSharingLinkTypeOption
                }
                Write-Verbose -Verbose -Message "Setting default sharing link type to $($defaultSharingLinkTypeOption)"
            }

            #Set the default link permission type 
            if ($eumDefaultLinkPermission){
                 switch ($eumDefaultLinkPermission){
                    'View' {$defaultLinkPermissionOption = "View" ; Break}
                    'Edit' {$defaultLinkPermissionOption = "Edit" ; Break}
                }
                Write-Verbose -Verbose -Message "Setting default link permission to $($defaultLinkPermissionOption)"
                Set-PnPSite -Identity $siteURL -DefaultLinkPermission $defaultLinkPermissionOption  
            }  

            # Set the site collection admin
            if ($SiteCollectionAdministrator -ne "") {
                Add-PnPSiteCollectionAdmin -Owners $SiteCollectionAdministrator -Connection $connSite
            }
            # If the Owner is different than the requestor, it is because the requestor is external
            # They can only be added as an admin if external sharing is enabled
            if (($requestor -ne $owner) -and ($externalSharingOption -ne "Disabled")) {
                Add-PnPSiteCollectionAdmin -Owners $requestor -Connection $connSite
            }

            # Add Everyone group if on-prem and Public
            if (-not $IsSharePointOnline -and $siteVisibility -eq "Public") {
                Set-PnPWebPermission -User "c:0(.s|true" -AddRole "Read" -Connection $connSite
            }

            # add the site to hub site, if it configured
            if ($IsSharePointOnline -and $parentHubSite -and $joinHubSite) {
                Write-Verbose -Verbose -Message "Adding the site ($($siteURL)) to the parent hub site($($parentURL))."
                Add-PnPHubSiteAssociation -Site $siteURL -HubSite $parentURL -Connection $connSite
            }

            if ($pnpSiteTemplate -ne "") {
                # intermittent error applying SiteTemplate before TaxCatchAllField has been allocated https://github.com/pnp/PnP-PowerShell/issues/1180
                # suggested workaround is a loop to check for existence of that field prior to applying template
                # Let OOTB taxonomy TaxCatchAllField column to be provisioned. This column is a depencency we cannot skip
                $retries = 0
                $TaxCatchAllField = $null
                $TaxCatchAllField =  Get-PnPField -Identity "f3b0adf9-c1a2-4b02-920d-943fba4b3611" -ErrorAction SilentlyContinue
                while(($retries -lt 36) -and ($null -eq $TaxCatchAllField)){
                    $retries += 1
                    Write-Verbose -Verbose -Message "Waiting for TaxCatchAllField column to be provisioned..." 
                    Start-Sleep -Seconds 5
                    $TaxCatchAllField =  Get-PnPField -Identity "f3b0adf9-c1a2-4b02-920d-943fba4b3611" -ErrorAction SilentlyContinue
                }

                $retries = 0
                $pnpTemplateApplied = $false
                while (($retries -lt 20) -and ($pnpTemplateApplied -eq $false)) {
                    Write-Verbose -Verbose -Message "Applying template $($pnpSiteTemplate) Please wait..."
                    try {
                        $retries += 1
                        Set-PnPTraceLog -On -Level Debug
                        Invoke-PnPSiteTemplate -Path $pnpSiteTemplate -Connection $connSite
                        $pnpTemplateApplied = $true
                    }
                    catch {      
                        Write-Verbose -Verbose -Message "Failed applying PnP template."
                        Write-Verbose -Verbose -Message $_
                        Start-Sleep -Seconds 30
                    }
                }
            }
            
            # Check if a team was created
            if ($IsSharePointOnline -and $eumCreateTeam) {
                Write-Verbose -Verbose -Message "groupId = $($groupId), generalChannelId = $($generalChannelId)"

                if ($eumCreateOneNote) {
                    AddOneNoteTeamsChannelTab -groupId $groupId -channelName 'General' -teamsChannelId $generalChannelId -siteURL $siteURL -teamName $team.DisplayName
                }

                if ($eumCreatePlanner) {
                    $planId = AddTeamPlanner -groupId $groupId -planTitle "$($siteTitle) Planner"
                    AddPlannerTeamsChannelTab -groupId $groupId -planTitle "$($siteTitle) Planner" -planId $planId -channelName 'General' -teamsChannelId $generalChannelId  
                }
            }

            # Set the breadcrumb HTML
            [string]$breadcrumbHTML = GetBreadcrumbHTML -siteURL $siteURL -siteTitle $siteTitle -parentURL $parentURL

            # Provision the Site Metadata list in the newly created site and add the entry
            $siteMetadataPnPTemplate = "$pnpTemplatePath\EUMSites.SiteMetadataList.xml"
            # Only do this if the template exists.  It is not required if security trimmed A-Z sites list is not needed
            if (Test-Path $siteMetadataPnPTemplate) {
                $connSite = Helper-Connect-PnPOnline -Url $siteURL
                $retries = 0
                $pnpTemplateApplied = $false
                while (($retries -lt 20) -and ($pnpTemplateApplied -eq $false)) {
                    Write-Verbose -Verbose -Message "Importing Site Metadata list with PnPTemplate $($siteMetadataPnPTemplate)"
                    try {
                        $retries += 1
                        Invoke-PnPSiteTemplate -Path $siteMetadataPnPTemplate -Connection $connSite

                        [hashtable]$newListItemValues = @{ }

                        $newListItemValues.Add("Title", $siteTitle)
                        $newListItemValues.Add("EUMAlias", $alias)
                        $newListItemValues.Add("EUMDivision", $Division)
                        $newListItemValues.Add("EUMGroupSummary", $groupSummary)
                        $newListItemValues.Add("EUMParentURL", $parentURL)
                        $newListItemValues.Add("SitePurpose", $sitePurpose)
                        $newListItemValues.Add("EUMSiteTemplate", $eumSiteTemplate)
                        $newListItemValues.Add("EUMSiteURL", $siteURL)
                        $newListItemValues.Add("EUMSiteVisibility", $siteVisibility)
                        $newListItemValues.Add("EUMSiteCreated", [System.DateTime]::Now)
                        $newListItemValues.Add("EUMIsSubsite", $false)
                        $newListItemValues.Add("EUMBreadcrumbHTML", $breadcrumbHTML)

                        [Microsoft.SharePoint.Client.ListItem]$spListItem = Add-PnPListItem -List "Site Metadata" -Values $newListItemValues -Connection $connSite
                        $pnpTemplateApplied = $true
                    }
                    catch {      
                        Write-Verbose -Verbose -Message "Failed applying PnP template."
                        Write-Verbose -Verbose -Message $_
                        Start-Sleep -Seconds 30
                    }
                }
            }
            
            # Reconnect to the master site and update the site collection list
            $connLandingSite = Helper-Connect-PnPOnline -Url $SiteCollectionFullURL

            # Set the breadcrumb and site URL
            [Microsoft.SharePoint.Client.ListItem]$spListItem = Set-PnPListItem -List $SiteListName -Identity $pendingSite.Id -Values @{ "EUMBreadcrumbHTML" = $breadcrumbHTML; "EUMSiteURL" = $siteURL; "EUMParentURL" = $parentURL } -Connection $connLandingSite
        }
    }
    else {
        Write-Verbose -Verbose -Message "No sites pending creation"
    }

    return $True
}

function CreateTeamChannel () {
    Param
    (
        [Parameter (Mandatory = $true)][int]$listItemID
    )

    try {
        Write-Verbose -Verbose -Message "Retrieving teams channel request details for listItemID $($listItemID)..."
        $connLandingSite = Helper-Connect-PnPOnline -Url $SiteCollectionFullURL
        $channelDetails = Get-PnPListItem -List $TeamsChannelsListName -Id $listItemID -Fields "ID", "Title", "IsPrivate", "Description", "TeamSiteURL", "Description", "CreateOneNoteSection", "CreateChannelPlanner", "ChannelTemplate" -Connection $connLandingSite

        [string]$channelName = $channelDetails["Title"]
        [boolean]$isPrivate = $channelDetails["IsPrivate"]
        [string]$siteURL = $channelDetails["TeamSiteURL"]
        [string]$channelDescription = $channelDetails["Description"]
        [boolean]$createOneNote = $channelDetails["CreateOneNoteSection"]
        [boolean]$createPlanner = $channelDetails["CreateChannelPlanner"]
        [string]$channelTemplateId = $channelDetails["ChannelTemplate"].LookupId

        Disconnect-PnPOnline

        # Get the Office 365 Group ID
        Write-Verbose -Verbose -Message "Retrieving group ID for site $($siteURL)..."
        $connAdmin = Helper-Connect-PnPOnline -Url $AdminURL
        $spSite = Get-PnPTenantSite -Url $siteURL -Connection $connAdmin
        $groupId = $spSite.GroupId
        Disconnect-PnPOnline
    }
    catch {
        Write-Error "Failed retrieving information for listItemID $($listItemID)"
        Write-Error $_
        return $false    
    }


    try {
        # Create the new channel in Teams
        Write-Verbose -Verbose -Message "Creating channel $($channelName)..."
        Helper-Connect-MicrosoftTeams -Credential $SPCredentials
        $teamsChannel = New-TeamChannel -GroupId $groupId -DisplayName $channelName -Description $channelDescription
        $teamsChannelId = $teamsChannel.Id
        Disconnect-MicrosoftTeams

        if ($createOneNote) {
            Write-Verbose -Verbose -Message "Configuring OneNote for $($channelName)..."
            $team = Get-Team -groupId $groupId
            $teamName = $team.DisplayName
            AddOneNoteTeamsChannelTab -groupId $groupId -channelName $channelName -teamsChannelId $teamsChannelId -siteURL $siteURL -teamName $teamName
        }

        if ($createPlanner) {
            Write-Verbose -Verbose -Message "Creating Planner for $($channelName)..."
            $planId = AddTeamPlanner -groupId $groupId -planTitle "$($channelName) Planner"
            AddPlannerTeamsChannelTab -groupId $groupId -planTitle "$($channelName) Planner" -planId $planId -channelName $channelName -teamsChannelId $teamsChannelId          
        }

        # Apply implementation specific customizations
        ApplyChannelCustomizations -listItemID $listItemID

        # update the SP list with the ChannelCreationDate
        Write-Verbose -Verbose -Message "Updating ChannelCreationDate..."

        $connLandingSite = Helper-Connect-PnPOnline -Url $SiteCollectionFullURL

        $spListItem = Set-PnPListItem -List $TeamsChannelsListName -Identity $listItemID -Values @{"ChannelCreationDate" = (Get-Date) } -Connection $connLandingSite
        Disconnect-PnPOnline
    }
    catch {
        Write-Error "Failed creating teams channel $($channelName)"
        Write-Error $_
        return $false   
    }
}

function Check-RunbookLock {
    [String] $ServicePrincipalConnectionName = 'AzureRunAsConnection'
    $AutomationAccountName = Get-AutomationVariable -Name 'AutomationAccountName'
    $ResourceGroupName = Get-AutomationVariable -Name 'ResourceGroupName'
    $AutomationJobID = $PSPrivateMetadata.JobId.Guid

    Write-Verbose "Set-RunbookLock Job ID: $AutomationJobID"

    $ServicePrincipalConnection = Get-AutomationConnection -Name $ServicePrincipalConnectionName   
    if (!$ServicePrincipalConnection) {
        $ErrorString = 
        @"
        Service principal connection $ServicePrincipalConnectionName not found.  Make sure you have created it in Assets. 
        See http://aka.ms/runasaccount to learn more about creating Run As accounts. 
"@
        throw $ErrorString
    }  	
    
    Add-AzAccount `
        -ServicePrincipal `
        -Tenant $ServicePrincipalConnection.TenantId `
        -ApplicationId $ServicePrincipalConnection.ApplicationId `
        -CertificateThumbprint $ServicePrincipalConnection.CertificateThumbprint | Write-Verbose

    # Get the information for this job so we can retrieve the Runbook Id
    $CurrentJob = Get-AzAutomationJob -AutomationAccountName $AutomationAccountName -ResourceGroupName $ResourceGroupName -Id $AutomationJobID
    Write-Verbose "Set-RunbookLock AutomationAccountName: $($CurrentJob.AutomationAccountName)"
    Write-Verbose "Set-RunbookLock RunbookName: $($CurrentJob.RunbookName)"
    Write-Verbose "Set-RunbookLock ResourceGroupName: $($CurrentJob.ResourceGroupName)"
    
    $AllJobs = Get-AzAutomationJob -AutomationAccountName $CurrentJob.AutomationAccountName `
        -ResourceGroupName $CurrentJob.ResourceGroupName `
        -RunbookName $CurrentJob.RunbookName | Sort-Object -Property CreationTime, JobId | Select-Object -Last 10

    foreach ($job in $AllJobs) {
        Write-Verbose "JobID: $($job.JobId), CreationTime: $($job.CreationTime), Status: $($job.Status)"
    }

    $AllActiveJobs = Get-AzAutomationJob -AutomationAccountName $CurrentJob.AutomationAccountName `
        -ResourceGroupName $CurrentJob.ResourceGroupName `
        -RunbookName $CurrentJob.RunbookName | Where -FilterScript { ($_.Status -ne "Completed") `
            -and ($_.Status -ne "Failed") `
            -and ($_.Status -ne "Stopped") } 

    Write-Verbose "AllActiveJobs.Count $($AllActiveJobs.Count)"

    # If there are any active jobs for this runbook, return false. If this is the only job
    # running then return true
    If ($AllActiveJobs.Count -gt 1) {
        # In order to prevent a race condition (although still possible if two jobs were created at the 
        # exact same time), let this job continue if it is the oldest created running job
        $OldestJob = $AllActiveJobs | Sort-Object -Property CreationTime, JobId | Select-Object -First 1
        Write-Verbose "AutomationJobID: $($AutomationJobID), OldestJob.JobId: $($OldestJob.JobId)"

        # If this job is not the oldest created job we will suspend it and let the oldest one go through.
        # When the oldest job completes it will call Set-RunbookLock to make sure the next-oldest job for this runbook is resumed.
        if ($AutomationJobID -ne $OldestJob.JobId) {
            Write-Verbose "Returning false as there is an older currently running job for this runbook already"
            return $false
        }
        else {
            Write-Verbose "Returning true as this is the oldest currently running job for this runbook"
            return $true
        }
    }
    Else {
        Write-Verbose "No other currently running jobs for this runbook"
        return $true
    }
}

# Use this for debugging direct runs of this Runbook
# $testItemID = 82

if ($testItemID -ne -1) {
    LoadEnvironmentSettings
    ProvisionSite -listItemID $testItemID
}
