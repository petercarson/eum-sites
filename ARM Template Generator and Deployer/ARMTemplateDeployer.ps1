<#
    ARM Template Deployer Tool v1.6.0

    Prerequisite:   Azure Az PowerShell module must be installed.
                    https://docs.microsoft.com/en-us/powershell/azure/install-az-ps

    Tested:         Fully tested against Azure Az PowerShell 5.2.0

    Purpose:
    - Connect to $SubscriptionId
    - Check if Target Resource Group exists.  If not, create it.
    - Deploy ARM template
    - After the ARM template is fully deployed, dynamically retrieve the Automation Account from $ResourceGroupName
    - Dynamically attain the list of Runbook PowerShell files from $TargetFolder
    - Import each Runbook PowerShell file into the Automation Account
    - If Runbook already exists in the Automation Account, it will do an update.
    - Publish each Runbook
    - Provision "Run As Account"
    - Install Automation Account modules

    NOTE:
    - The ARM template is responsible for provisioning the Azure Automation Account.
#>

# Define your set of Automation Account Modules in the object below.
# If you need to install a specific version of the module, define it in the associated ModuleVersion.
# If ModuleVerison is blank, the latest version of the module will be installed.
# Modules will be installed in the order listed in the object.
$objModules = @(
    [PSCustomObject]@{ModuleName = "Az.Accounts"; ModuleVersion = "" }
    [PSCustomObject]@{ModuleName = "Az.Resources"; ModuleVersion = "" }
    [PSCustomObject]@{ModuleName = "AzureAD"; ModuleVersion = "" }
    [PSCustomObject]@{ModuleName = "Microsoft.Online.SharePoint.PowerShell"; ModuleVersion = "" }
    [PSCustomObject]@{ModuleName = "MicrosoftTeams"; ModuleVersion = "" }
    [PSCustomObject]@{ModuleName = "PnP.PowerShell"; ModuleVersion = "" }
)

function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string]$Message,
 
        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [ValidateSet('Information','Warning','Error')]
        [string]$Severity = 'Information',

        [Parameter()]
        [Boolean]$WriteToLogFile = $true
    )

    <#
        .SYNOPSIS
            Write-Log

        .DESCRIPTION
            Use Write-Log to write a timestamped message in a log file in CSV format.
            This is useful for troubleshooting.  The message will be echoed via Write-Host at the same time.

        .PARAMETER Message
            The message to be logged.

        .PARAMETER Severity
            Level of severity: Information, Warning, Error

        .PARAMETER WriteToLogFile
            Flag to determine whether to write to the log file. Default is True.
    #>
 
    $TimeStamp = (Get-Date -f "yyyy-MM-dd hh:mm:ss tt")
    $path = Get-Location
    $scriptName = [io.path]::GetFileNameWithoutExtension($MyInvocation.PSCommandPath)

    if ($WriteToLogFile) {
        [pscustomobject]@{
            Time = $TimeStamp
            Message = $Message
            Severity = $Severity
        } | Export-Csv -Path "$path\$scriptName.log" -Append -NoTypeInformation
    }

    Write-Host "[$TimeStamp] $Message"
}

function ImportRunbook {
    Param (
        [Parameter(Mandatory = $true)][string] $ResourceGroupName,      # Target Resource Group
        [Parameter(Mandatory = $true)][string] $AutomationAccountName,  # Target Automation Account
        [Parameter(Mandatory = $true)][string] $RunbookName,            # Name of Runbook PowerShell
        [Parameter(Mandatory = $true)][string] $TargetFolder            # Local target folder where the Runbook PowerShell resides
    )

    <#
        .SYNOPSIS
            ImportRunbook

        .DESCRIPTION
            Import and Publish an Azure Automation Runbook.
            If the Runbook already exists, it will update the Runbook.

        .PARAMETER ResourceGroupName
            Target Resource Group

        .PARAMETER AutomationAccountName
            Target Automation Account

        .PARAMETER RunbookName
            Name of Runbook PowerShell

        .PARAMETER TargetFolder
            Local target folder where the Runbook PowerShell resides
    #>

    <#
        Supported Runbook Types:

        - PowerShell
        - GraphicalPowerShell
        - PowerShellWorkflow
        - GraphicalPowerShellWorkflow
        - Graph - The value Graph is obsolete. It is equivalent to GraphicalPowerShellWorkflow.
        - Python2 - Current supported Python version.  Update this to support other Python versions in the future.

        See: https://docs.microsoft.com/en-us/azure/automation/manage-runbooks#import-a-runbook
    #>

    $fileExtension = $RunbookName.Split(".")[1]
    switch ($fileExtension) {
        "ps1" {
            $runbookType = "PowerShell"
        }
        "py" {
            $runbookType = "Python2" # current supported Python version. Might need to update this in the future.
        }
        "graphrunbook" {
            $runbookType = "GraphicalPowerShell"
        }
        Default {
            $runbookType = "PowerShell"
        }
    }

    # Import and Publish the Automation Runbook
    Write-Log "Importing and publishing $runbookType Runbook '$TargetFolder\$RunbookName' in '$AutomationAccountName' Automation Account..."
    Import-AzAutomationRunbook -Name $RunbookName -Path "$TargetFolder\$RunbookName" -ResourceGroupName $ResourceGroupName -AutomationAccountName $AutomationAccountName -Type $runbookType -Published -Force
}

function InstallAutomationAccountModules{
    param(
        [Parameter(Mandatory = $true)][string] $ResourceGroupName,
        [Parameter(Mandatory = $true)][string] $AutomationAccountName,
        [Parameter(Mandatory = $true)][Object[]] $objModules
    )

    <#
        .SYNOPSIS
        InstallAutomationAccountModules

        .DESCRIPTION
        Install Automation Account Modules.
    #>

    $objModules | ForEach-Object {
        $ModuleName = $_.ModuleName
        $ModuleVersion = $_.ModuleVersion
        Write-Log "Installing Module Name = $ModuleName $ModuleVersion ..."
        New-AzAutomationModule -AutomationAccountName $AutomationAccountName -ResourceGroupName $ResourceGroupName -Name $ModuleName -ContentLinkUri "https://www.powershellgallery.com/api/v2/package/$ModuleName/$ModuleVersion"
        
        $Retries = 0;
        # Sleep for a few seconds to allow the module to become active (ordinarily takes a few seconds)
        Start-Sleep -s 30
        $ModuleReady = (Get-AzAutomationModule -AutomationAccountName $AutomationAccountName -Name $ModuleName -ResourceGroupName $ResourceGroupName -ErrorAction SilentlyContinue).ProvisioningState

        # If new module is not ready, retry.
        While ($ModuleReady -ne "Succeeded" -and $Retries -le 6) {
            $Retries++;
            Write-Log "$ModuleName - Retry $Retries..."
            Start-Sleep -s 20
            $ModuleReady = (Get-AzAutomationModule -AutomationAccountName $AutomationAccountName -Name $ModuleName -ResourceGroupName $ResourceGroupName -ErrorAction SilentlyContinue).ProvisioningState
        }
    }
}

function ProvisionRunAsAccount 
{
    Param (
        [Parameter(Mandatory = $true)]
        [String] $ResourceGroupName,

        [Parameter(Mandatory = $true)]
        [String] $AutomationAccountName,

        [Parameter(Mandatory = $true)]
        [String] $ApplicationDisplayName,

        [Parameter(Mandatory = $true)]
        [String] $SubscriptionId,

        [Parameter(Mandatory = $true)]
        [String] $SelfSignedCertPlainPassword,

        [Parameter(Mandatory = $false)]
        [string] $EnterpriseCertPathForRunAsAccount,

        [Parameter(Mandatory = $false)]
        [String] $EnterpriseCertPlainPasswordForRunAsAccount,

        [Parameter(Mandatory = $false)]
        [String] $EnterpriseCertPathForClassicRunAsAccount,

        [Parameter(Mandatory = $false)]
        [int] $SelfSignedCertNoOfMonthsUntilExpired = 24
    )

    <#
        .SYNOPSIS
        ProvisionRunAsAccount

        .DESCRIPTION
        Provision Run As Account.
    #>

    function CreateSelfSignedCertificate
    {
        Param (
            [string] $certificateName, 
            [string] $selfSignedCertPlainPassword,
            [string] $certPath, 
            [string] $certPathCer, 
            [int] $selfSignedCertNoOfMonthsUntilExpired 
        ) 

        $Cert = New-SelfSignedCertificate -DnsName $certificateName -CertStoreLocation cert:\LocalMachine\My -KeyExportPolicy Exportable -Provider "Microsoft Enhanced RSA and AES Cryptographic Provider" -NotAfter (Get-Date).AddMonths($selfSignedCertNoOfMonthsUntilExpired) -HashAlgorithm SHA256
        $CertPassword = ConvertTo-SecureString $selfSignedCertPlainPassword -AsPlainText -Force

        Export-PfxCertificate -Cert ("Cert:\localmachine\my\" + $Cert.Thumbprint) -FilePath $certPath -Password $CertPassword -Force | Write-Verbose
        Export-Certificate -Cert ("Cert:\localmachine\my\" + $Cert.Thumbprint) -FilePath $certPathCer -Type CERT | Write-Verbose
    }

    function CreateServicePrincipal
    {
        Param (
            [System.Security.Cryptography.X509Certificates.X509Certificate2] $PfxCert, 
            [string] $ApplicationDisplayName
        )

        $keyValue = [System.Convert]::ToBase64String($PfxCert.GetRawCertData())
        $keyId = (New-Guid).Guid

        # Create an Azure AD application, AD App Credential, AD ServicePrincipal
        # Requires Application Developer Role, but works with Application administrator or GLOBAL ADMIN
        $Application = New-AzADApplication -DisplayName $ApplicationDisplayName -HomePage ("http://" + $ApplicationDisplayName) -IdentifierUris ("http://" + $keyId)

        # Requires Application administrator or GLOBAL ADMIN
        $ApplicationCredential = New-AzADAppCredential -ApplicationId $Application.ApplicationId -CertValue $keyValue -StartDate $PfxCert.NotBefore -EndDate $PfxCert.NotAfter

        # Requires Application administrator or GLOBAL ADMIN
        $ServicePrincipal = New-AzADServicePrincipal -ApplicationId $Application.ApplicationId
        $GetServicePrincipal = Get-AzADServicePrincipal -ObjectId $ServicePrincipal.Id

        # Sleep for a few seconds to allow the service principal application to become active (ordinarily takes a few seconds)
        Start-Sleep -s 15

        # Requires User Access Administrator or Owner.
        $NewRole = New-AzRoleAssignment -RoleDefinitionName Contributor -ServicePrincipalName $Application.ApplicationId -ErrorAction SilentlyContinue
        $Retries = 0;
        # Sleep for a few seconds to allow the new role assignment to become active (ordinarily takes a few seconds)
        Start-Sleep -s 25
        $NewRole = Get-AzRoleAssignment -ServicePrincipalName $Application.ApplicationId -ErrorAction SilentlyContinue

        # If new role assignment is not ready, retry.
        While ($null -eq $NewRole -and $Retries -le 6) {
            New-AzRoleAssignment -RoleDefinitionName Contributor -ServicePrincipalName $Application.ApplicationId | Write-Verbose -ErrorAction SilentlyContinue
            Start-Sleep -s 10
            $NewRole = Get-AzRoleAssignment -ServicePrincipalName $Application.ApplicationId -ErrorAction SilentlyContinue
            $Retries++;
        }

        return $Application.ApplicationId.ToString();
    }

    function CreateAutomationCertificateAsset 
    {
        Param (
            [string] $ResourceGroupName, 
            [string] $AutomationAccountName, 
            [string] $CertifcateAssetName, 
            [string] $CertPath, 
            [string] $CertPlainPassword, 
            [Boolean] $Exportable
        )

        $CertPassword = ConvertTo-SecureString $certPlainPassword -AsPlainText -Force

        Remove-AzAutomationCertificate -ResourceGroupName $ResourceGroupName -AutomationAccountName $AutomationAccountName -Name $CertifcateAssetName -ErrorAction SilentlyContinue
        New-AzAutomationCertificate -ResourceGroupName $ResourceGroupName -AutomationAccountName $AutomationAccountName -Path $CertPath -Name $CertifcateAssetName -Password $CertPassword -Exportable:$Exportable | write-verbose
    }

    function CreateAutomationConnectionAsset 
    {
        Param (
            [string] $ResourceGroupName, 
            [string] $AutomationAccountName, 
            [string] $ConnectionAssetName, 
            [string] $ConnectionTypeName, 
            [System.Collections.Hashtable] $ConnectionFieldValues 
        )

        Remove-AzAutomationConnection -ResourceGroupName $ResourceGroupName -AutomationAccountName $AutomationAccountName -Name $ConnectionAssetName -Force -ErrorAction SilentlyContinue
        New-AzAutomationConnection -ResourceGroupName $ResourceGroupName -AutomationAccountName $AutomationAccountName -Name $ConnectionAssetName -ConnectionTypeName $ConnectionTypeName -ConnectionFieldValues $ConnectionFieldValues
    }	

    ##################
    # Main Execution #
    ##################

    # Create a Run As account by using a service principal
    $CertifcateAssetName = "AzureRunAsCertificate"
    $ConnectionAssetName = "AzureRunAsConnection"
    $ConnectionTypeName = "AzureServicePrincipal"

    if ($EnterpriseCertPathForRunAsAccount -and $EnterpriseCertPlainPasswordForRunAsAccount) {
        $PfxCertPathForRunAsAccount = $EnterpriseCertPathForRunAsAccount
        $PfxCertPlainPasswordForRunAsAccount = $EnterpriseCertPlainPasswordForRunAsAccount
    }
    else {
        $CertificateName = $AutomationAccountName + $CertifcateAssetName
        $PfxCertPathForRunAsAccount = Join-Path $env:TEMP ($CertificateName + ".pfx")
        $PfxCertPlainPasswordForRunAsAccount = $SelfSignedCertPlainPassword
        $CerCertPathForRunAsAccount = Join-Path $env:TEMP ($CertificateName + ".cer")
        CreateSelfSignedCertificate $CertificateName $PfxCertPlainPasswordForRunAsAccount $PfxCertPathForRunAsAccount $CerCertPathForRunAsAccount  $SelfSignedCertNoOfMonthsUntilExpired
    }

    # Create a service principal
    $PfxCert = New-Object -TypeName System.Security.Cryptography.X509Certificates.X509Certificate2 -ArgumentList @($PfxCertPathForRunAsAccount, $PfxCertPlainPasswordForRunAsAccount)
    $ApplicationId = CreateServicePrincipal -PfxCert $PfxCert -ApplicationDisplayName $ApplicationDisplayName

    # Create the Automation certificate asset
    CreateAutomationCertificateAsset -ResourceGroupName $ResourceGroupName -AutomationAccountName $AutomationAccountName -CertifcateAssetName $CertifcateAssetName -CertPath $PfxCertPathForRunAsAccount -CertPlainPassword $PfxCertPlainPasswordForRunAsAccount -Exportable $true

    # Populate the ConnectionFieldValues
    $SubscriptionInfo = Get-AzSubscription -SubscriptionId $SubscriptionId
    $TenantID = $SubscriptionInfo | Select-Object TenantId -First 1
    $Thumbprint = $PfxCert.Thumbprint
    $ConnectionFieldValues = @{"ApplicationId" = $ApplicationId; "TenantId" = $TenantID.TenantId; "CertificateThumbprint" = $Thumbprint; "SubscriptionId" = $SubscriptionId }

    # Create an Automation connection asset named AzureRunAsConnection in the Automation account. This connection uses the service principal.
    CreateAutomationConnectionAsset -ResourceGroupName $ResourceGroupName -AutomationAccountName $AutomationAccountName -ConnectionAssetName $ConnectionAssetName -ConnectionTypeName $ConnectionTypeName -ConnectionFieldValues $ConnectionFieldValues
}

function DeployARMTemplate {
    Param (
        [Parameter(Mandatory = $true)][string]  $TenantId,
        [Parameter(Mandatory = $true)][string]  $SubscriptionId,
        [Parameter(Mandatory = $true)][string]  $ResourceGroupName,
        [Parameter(Mandatory = $true)][string]  $Location,
        [Parameter(Mandatory = $true)][string]  $TargetFolder,
        [Parameter(Mandatory = $false)][String] $TemplateFilename = "MainTemplate.json",
        [Parameter(Mandatory = $false)][String] $TemplateParameterFilename = "MainTemplate.parameters.json"
    )

    <#
        .SYNOPSIS
        Deploy ARM Template to an Azure Resource Group

        .DESCRIPTION
        Deploy the ARM Template, optional paramter file, and all associated Runbooks to a target Azure Resource Group.

        .PARAMETER SubscriptionId
        Target Azure SubscriptionId.

        .PARAMETER ResourceGroupName
        Target Resource Group.

        .PARAMETER Location
        Location of Resource Group.

        .PARAMETER TargetFolder
        Local target folder where the ARM Template, optional parameter file, and Runbooks reside.

        .PARAMETER TemplateFilename
        Template filename (default to "MainTemplate.json").

        .PARAMETER TemplateParameterFilename
        Optional: Template parameter filename (default to "MainTemplate.parameters.json").  This will be used if the file exists.
    #>    

    function GenerateRandomPassword {
        Param (
            [Parameter()]
            [int]$MinimumPasswordLength = 5,
            [Parameter()]
            [int]$MaximumPasswordLength = 10,
            [Parameter()]
            [switch]$ConvertToSecureString
        )
        
        Add-Type -AssemblyName 'System.Web'
        $length = Get-Random -Minimum $MinimumPasswordLength -Maximum $MaximumPasswordLength
        $password = [System.Web.Security.Membership]::GeneratePassword($length, 1)
        if ($ConvertToSecureString.IsPresent) {
            ConvertTo-SecureString -String $password -AsPlainText -Force
        } else {
            $password
        }
    }

    # Connect to Azure
    Connect-AzAccount -Tenant $TenantId -SubscriptionId $SubscriptionId

    # Make sure $SubscriptionId is a valid Azure Subscription in this tenant.
    if ( Get-AzSubscription -SubscriptionId $SubscriptionId ) {
        # Subscription is valid.  We can proceed.
        Set-AzContext -SubscriptionId $SubscriptionId

        # Check if $ResourceGroupName exists in the selected Azure Subscription.
        Get-AzResourceGroup -Name $ResourceGroupName -ErrorVariable notFound -ErrorAction SilentlyContinue
    
        if( $notFound ) {
            # If resource group does not exist, create it!
            Write-Log "Creating resource group '$ResourceGroupName'..."
            New-AzResourceGroup -Name $ResourceGroupName -Location $Location
        } else {
            Write-Log "Updating resource group '$ResourceGroupName'..."
        }

        # Deploy the full ARM Template in $ResourceGroupName.
        # This is so that the Automation Account(s) would be provisioned before we import the Runbook PowerShell scripts.
        If ( Test-Path "$TargetFolder\$TemplateParameterFilename" ) {
            Write-Log "Template parameter file '$TargetFolder\$TemplateParameterFilename' provided..."
            New-AzResourceGroupDeployment -ResourceGroupName $ResourceGroupName -TemplateFile "$TargetFolder\$TemplateFilename" -TemplateParameterFile "$TargetFolder\$TemplateParameterFilename"
        } else {
            # Template parameter file not available
            New-AzResourceGroupDeployment -ResourceGroupName $ResourceGroupName -TemplateFile "$TargetFolder\$TemplateFilename"
        }

        # After deploying ARM Template...
        Write-Log "DONE: ARM Deployment"

        # Test to see if Automation Account is created...
        # Get all Automation Accounts for $ResourceGroupName.
        $objAutomationAccounts = Get-AzAutomationAccount -ResourceGroupName $ResourceGroupName | Select-Object AutomationAccountName

        # Dynamically attain the list of Runbook PowerShell files from $TargetFolder
        $listRunbooks = (Get-ChildItem $TargetFolder\* -include *.ps1, *.graphrunbook, *.py).Name

        # For each Automation Account, get a list of all Runbooks for that specific Automation Account.
        foreach( $AutomationAccountSource in $objAutomationAccounts ) {
            Write-Log "Automation account '$($AutomationAccountSource.AutomationAccountName)' found!"
            $objRunbooks = Get-AzAutomationRunbook -ResourceGroupName $ResourceGroupName -AutomationAccountName $AutomationAccountSource.AutomationAccountName | Select-Object Name

            foreach( $RunbookName in $listRunbooks ) {
                # Create Runbook if it does not already exist in Resource Group.  Otherwise, update Runbook if it already exists.
                if( $objRunbooks -match $RunbookName ) {
                    # Runbook already exists.
                    Write-Log "Runbook '$RunbookName' already exists.  Update Runbook."
                }
                ImportRunbook -ResourceGroupName $ResourceGroupName -AutomationAccountName $AutomationAccountSource.AutomationAccountName -RunbookName $RunbookName -TargetFolder "$TargetFolder\Runbooks"
            }

            # Provision Run As Account
            Write-Log "Provisioning Run As Account in Automation Account '$($AutomationAccountSource.AutomationAccountName)'..."
            $SelfSignedCertPlainPassword = GenerateRandomPassword
            ProvisionRunAsAccount -ResourceGroupName $ResourceGroupName -AutomationAccountName $AutomationAccountSource.AutomationAccountName -SubscriptionId $SubscriptionId -ApplicationDisplayName $AutomationAccountSource.AutomationAccountName -SelfSignedCertPlainPassword $SelfSignedCertPlainPassword

            Write-Log "*** Installing Automation Account Modules..."
            InstallAutomationAccountModules -ResourceGroupName $ResourceGroupName -AutomationAccountName $AutomationAccountSource.AutomationAccountName -objModules $objModules
        }
    }
    else {
        # Error handling: An invalid $SubscriptionId was provided
        Write-Log "Sorry, SubscriptionId '$SubscriptionId' was not found.  Please verify that the subscription exists and try again."
    }
}

function ReadConfiguration {
    Param (
        [Parameter(Mandatory = $false)][String] $ConfigFile = "config.json" 
    )

    <#
        .SYNOPSIS
        Read configuration data

        .PARAMETER ConfigFilename
        JSON file containing the Parameter Default Value Mapping.
    #>

    $global:config = $null

    if (Test-Path -Path $ConfigFile -ErrorAction SilentlyContinue) {
        $global:config = Get-Content -Path $ConfigFile -Raw | ConvertFrom-Json
    }
    else {
        $path = "$PSScriptRoot\$ConfigFile"
        if (Test-Path -Path $path -ErrorAction SilentlyContinue) {
            $global:config = Get-Content -Path $path -Raw | ConvertFrom-Json
        }
    }

    if ($global:config -eq $null) {
        Write-Log -Severity Error -Message "Exiting - failed to load configuration $ConfigFile"
        Exit
    }
}


################
# Unit-testing #
################

# Declare varaibles
#$SubscriptionId             = "b906693c-5167-4a56-bbf3-b56142ee7219"            # Target Azure Subscription
#$ResourceGroupName          = "EUMSites_Tom"                                    # Target Resource Group Name
#$Location                   = "Canada Central"                                  # Resource Group Location
#$TargetFolder               = "$PSScriptRoot\CreateSiteARMTemplate"             # Local target folder where the exported ARM template and Runbooks are stored.
#$TemplateFilename           = "template.json"
#$TemplateParameterFilename  = "parametersFile.json"                             # Option template parameter file. This will be used if the file exists.

# Tom's Subscription setings https://portal.azure.com/#@tabbottenvisionit.onmicrosoft.com/resource/subscriptions/aabfda3e-81d9-4d65-b7df-9b24020092ec/resourceGroups/ARMTestTeamsProvisioning/overview
#$SubscriptionId             = "aabfda3e-81d9-4d65-b7df-9b24020092ec"            # Target Azure Subscription - Subscription ID from Resource Group overview page
#$TenantId                   = "dff4b03f-2bee-4fd2-a5a2-160906fd80ab"            # Target Azure Subscription - AAD Tenant ID
#$ResourceGroupName          = "ARMTestTeamsProvisioning"                        # Target Resource Group Name
#$Location                   = "Canada Central"                                  # Resource Group Location
#$TargetFolder               = "$PSScriptRoot\CreateSiteARMTemplate"             # Local target folder where the exported ARM template and Runbooks are stored.
#$TemplateFilename           = "template.json"
#$TemplateParameterFilename  = "parametersFile.json"                             # Option template parameter file. This will be used if the file exists.

Write-Log "Start!"

$configFileInput = Read-Host "Enter the config.json full path (default - $configFile) "
if ($configFileInput -ne "") {
    $configFile = $configFileInput
}

ReadConfiguration -ConfigFile $configFile

$ResourceGroupName = Read-Host "Enter the resource group name to deploy to "

# Test execution
DeployARMTemplate -TenantId $config.TenantId -SubscriptionId $config.SubscriptionId -ResourceGroupName $ResourceGroupName -Location $config.Location -TargetFolder $config.TargetFolder `
                 -TemplateFilename $config.TemplateFilename -TemplateParameterFilename $config.TemplateParameterFilename

Write-Log "Done!"