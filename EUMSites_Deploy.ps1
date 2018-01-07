$DistributionFolder = (Split-Path $MyInvocation.MyCommand.Path)
. $DistributionFolder\EUMSites_Helper.ps1
[xml]$config = Get-Content -Path "$DistributionFolder\EUMSites.config"

#-----------------------------------------------------------------------
# Prompt for the environment defined in the config
#-----------------------------------------------------------------------
Write-Host "`n***** AVAILABLE ENVIRONMENTS *****" -ForegroundColor DarkGray
[string]$exampleEnvName
$config.settings.environments.environment | ForEach {
	Write-Host "$($_.name) - $($_.webApp.adminSiteURL)"
	$exampleEnvName = "$($_.name)"
	}
Write-Host "`n*********************************" -ForegroundColor DarkGray
[string]$EnvironmentName = Read-Host "Enter the Name of the environment from the above list, for example,"$exampleEnvName

[System.Xml.XmlLinkedNode]$Global:environment = $config.settings.environments.environment | Where { $_.name -eq $EnvironmentName }
#-----------------------------------------------------------------------
# Set variables based on environment selected
#-----------------------------------------------------------------------
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
	# Deploys the Application
	#-----------------------------------------------------------------------
	Write-Host "***** Deploying EUM Sites Application *****" -ForegroundColor DarkGray
	$DeploymentType = Read-Host "Choose the target for deployment: 1 for SharePoint, 2 for Azure, 3 for both" 
	if($DeploymentType -ne "1" -and $DeploymentType -ne "2" -and $DeploymentType -ne "3")
	{
		Write-Host "Deployment type has not been specified. Please choose the target for deployment: 1 for SharePoint, 2 for Azure, 3 for both."-ForegroundColor Red
	}

	if($DeploymentType -eq "1" -or $DeploymentType -eq "3") {
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
			}
			Write-Host "Connecting to "$SitesListSiteURL
			Connect-PnPOnline -Url $SitesListSiteURL -Credentials $SPCredentials
			Write-Host "Applying the EUM Sites Template to "$SitesListSiteURL
			Apply-PnPProvisioningTemplate -Path "$DistributionFolder\EUMSites.DeployTemplate.xml"
			Disconnect-PnPOnline
		}
		else
		{
			Write-Host "Required Windows Credential Manager 2.0 PowerShell Module not found. Please install the module by entering the following command in PowerShell: ""Install-Module -Name ""CredentialManager"" -RequiredVersion 2.0"""
			break
		}
	}

	if($DeploymentType -eq "2" -or $DeploymentType -eq "3") {
		#-----------------------------------------------------------------------
		# Azure Deployment
		#-----------------------------------------------------------------------
		#..coming soon
	}

 

 
 

