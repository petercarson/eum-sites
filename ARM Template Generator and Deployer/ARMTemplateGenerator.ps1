<#
    ARM Template Generator Tool v2.4.0

    Prerequisite:   Azure Az PowerShell module must be installed.
                    https://docs.microsoft.com/en-us/powershell/azure/install-az-ps

                        Install-Module Az
                        Install-Module Az.Resources
                        Install-Module Az.Automation

    Tested:         Fully tested against Azure Az PowerShell 5.2.0

    Purpose:        The ARM Template Generator Tool does the following:
                    - Connect to $SubscriptionId
                    - Get all Automation Accounts from $ResourceGroupName
                    - Dynamically retrieve the list of available Runbooks for each Automation Account
                    - Download runbooks from each Automation Account to $TargetFolder
                    - Export ARM Template from $ResourceGroupName to $TargetFolder\MainTemplate.json

                    After the export is done, the ARM Template Parameterization Tool will kick in.

                    The ARM Template Parameterization Tool is an automation tool that does the following:
                    - Find each of the workflow parameters defined in each of the Logic Apps,
                      and add ARM Template parameters to match these
                    - Replace workflow parameter default values with references to the ARM Template parameters
                    - Replace references to the subscription GUID and resource group location with built-in
                      ARM Template functions
                    - Automate the removal of extraneous parameters
                    - Automate the removal of extraneous resource objects
                    - Use config.json to define the parameter name and default value and aliases
                      This will automatically replace hard-coded values with dynamic parameters.

                    In addition, each resource in the resource group is exported individually and saved
                    in the SingleTemplates subfolder under the defined TargetFolder.
#>

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

function Export-SingleResources {
    Param (
        [Parameter(Mandatory = $true)][String]  $SubscriptionId,
        [Parameter(Mandatory = $true)][String]  $ResourceGroupName,
        [Parameter(Mandatory = $true)][String]  $TargetFolder,
        [Parameter(Mandatory = $false)][String]  $TargetSingleTemplatesFolder = "SingleTemplates"
    )

    <#
        .SYNOPSIS
        Export-SingleResources

        .DESCRIPTION
        Export individual resource ARM Templates inside a target Azure Resource Group.

        .PARAMETER SubscriptionId
        Source SubscriptionId.

        .PARAMETER ResourceGroupName
        Source Resource Group.

        .PARAMETER TargetFolder
        Local target folder to store exported ARM Template and Runbooks.
    #>

    Write-Log "Exporting individual resources..."

    $resources = Get-AzResource -ResourceGroupName $ResourceGroupName
    $NewTargetFolder = "$TargetFolder\$TargetSingleTemplatesFolder"

    # Check if $NewTargetFolder exists
    If ( !(Test-Path $NewTargetFolder) ) {
        New-Item -ItemType Directory -Force -Path $NewTargetFolder
    }

    foreach ($resource in $resources) {
        $resourceType = $resource.ResourceType
        $resourceName = $resource.Name
        Write-Log "Exporting individual ARM Template for $resourceName"
        Export-AzResourceGroup -ResourceGroupName $ResourceGroupName -Resource "/subscriptions/$SubscriptionId/resourceGroups/$ResourceGroupName/providers/$resourceType/$resourceName" -Path "$NewTargetFolder\$resourceName" -Force
    }
}

function ExportARMTemplateFromResourceGroup {
    Param (
        [Parameter(Mandatory = $true)][String]  $SubscriptionId,
        [Parameter(Mandatory = $true)][String]  $ResourceGroupName,
        [Parameter(Mandatory = $true)][String]  $TargetFolder,
        [Parameter(Mandatory = $false)][String] $TemplateFilename = "MainTemplate.json"
    )

    <#
        .SYNOPSIS
        Export ARM Template from Resource Group

        .DESCRIPTION
        Export the ARM Template and all associated Runbooks inside a target Azure Resource Group.

        .PARAMETER SubscriptionId
        Source SubscriptionId.

        .PARAMETER ResourceGroupName
        Source Resource Group.

        .PARAMETER TargetFolder
        Local target folder to store exported ARM Template and Runbooks.

        .PARAMETER TemplateFilename
        Template filename (default to "MainTemplate.json").
    #>

    # Connect to Azure
    Connect-AzAccount -Subscription $SubscriptionId

    # Make sure $SubscriptionId is a valid Azure Subscription in this tenant.
    if ( Get-AzSubscription -SubscriptionId $SubscriptionId ) {
        # Subscription is valid.  We can proceed.
        Set-AzContext -SubscriptionId $SubscriptionId

        # Get all resource groups for the selected Subscription.
        $objResourceGroups = Get-AzResourceGroup | Select-Object ResourceGroupName, Location
    
        # Make sure $ResourceGroupName is a valid Resource Group for the selected Azure Subscription in this tenant.
        if ( $objResourceGroups -match $ResourceGroupName ) {
            # Check if $TargetFolder exists.  If not, create it.
            If ( !(Test-Path $TargetFolder) ) {
                New-Item -ItemType Directory -Force -Path $TargetFolder
            }

            # Same for the Runbooks subfolder
            If ( !(Test-Path "$TargetFolder\Runbooks") ) {
                New-Item -ItemType Directory -Force -Path "$TargetFolder\Runbooks"
            }

            # Get all Automation Accounts for $ResourceGroupName.
            $objAutomationAccounts = Get-AzAutomationAccount -ResourceGroupName $ResourceGroupName | Select-Object AutomationAccountName
    
            # For each Automation Account, get a list of all Runbooks for that specific Automation Account.
            foreach ( $AutomationAccountNameSource in $objAutomationAccounts ) {
                $objRunbooks = Get-AzAutomationRunbook -ResourceGroupName $ResourceGroupName -AutomationAccountName $AutomationAccountNameSource.AutomationAccountName
    
                # For each Runbook, export the Azure Automation Runbook PowerShell source to $TargetFolder
                foreach ( $Runbook in $objRunbooks ) {
                    Export-AzAutomationRunbook -ResourceGroupName $ResourceGroupName -AutomationAccountName $AutomationAccountNameSource.AutomationAccountName -name $Runbook.Name -OutputFolder "$TargetFolder\Runbooks" -Force
                }
            }
    
            # Export the full ARM Template for $ResourceGroupName.
            Export-AzResourceGroup -ResourceGroupName $ResourceGroupName -Path "$TargetFolder\$TemplateFilename" -Force

            # Export each resource as its own template for $ResourceGroupName.
            Export-SingleResources -SubscriptionId $SubscriptionId -ResourceGroupName $ResourceGroupName -TargetFolder $TargetFolder
        }
        else {
            # Error handling: An invalid $ResourceGroupName was provided
            Write-Log "Sorry, resource group '$ResourceGroupName' was not found.  Please verify that the resource group exists and try again."
        }
    }
    else {
        # Error handling: An invalid $SubscriptionId was provided
        Write-Log "Sorry, SubscriptionId '$SubscriptionId' was not found.  Please verify that the subscription exists and try again."
    }
}

function ParameterizationPhase1 {
    Param (
        [Parameter(Mandatory = $true)][String] $TargetFolder,
        [Parameter(Mandatory = $true)][String] $TemplateFilename,
        [Parameter(Mandatory = $false)][object[]] $Parameters
    )

    <#
        .SYNOPSIS
        Parameterization Phase 1

        .DESCRIPTION
        Replace hard-coded parameters with dynamic counterparts.

        .PARAMETER TargetFolder
        Target folder where the ARN template resides.

        .PARAMETER TemplateFilename
        ARM Template filename.

        .PARAMETER Parameters
        Array of parameter objects
    #>

    [String] $TokenIdValue = '"id": "/subscriptions/'
    [String] $TokenAPI = "/managedApis"
    [String] $TokenLocation = '"location":'
    [String] $UpdatedARMTemplate = ""

    $TemplatePath = "$TargetFolder\$TemplateFilename"
    $armTemplate = Get-Content -Path $TemplatePath

    foreach ( $originalLine in $armTemplate ) {
        $revisedLine = $originalLine

        # Check if this line contains hard-coded tokens.
        if ( $originalLine.Contains($TokenIdValue)) {
            # If this line contains managedAPIs token, make it dynamic.
            if ( $originalLine.Contains($TokenAPI) ) {
                $index = $originalLine.LastIndexOf($TokenAPI)
                $ToPreserve = $originalLine.SubString($index)
                $ToPreserve = $ToPreserve.Substring(0, $ToPreserve.Length - 1)
                $revisedLine = """id"": ""[concat('/subscriptions/', subscription().subscriptionId, '/providers/Microsoft.Web/locations/', resourceGroup().location, '" + $ToPreserve + "')]"""
            }
        }

        # Make location dynamic: resourceGroup().location
        if ( $originalLine.Contains($TokenLocation)) {
            $index = $originalLine.LastIndexOf("""")
            $ToPreserve = $originalLine.SubString($index)
            $revisedLine = """location"": ""[resourceGroup().location]$ToPreserve"
        }

        # Replace any hard-coded defaultValue with dynamic "[parameters('...')]"
#  Targeted replacements of AutomationVariables, Logic App Params and Connection DisplayNames replaces this textwise global search/replace
#        if ( $parameters ) {
#            $parameters | ForEach-Object {
#                $QuotedValue = '"' + $_.paramValue + '"'
#                if ( $originalLine.Contains($QuotedValue)) {
#                    $dynamicParameter = "[parameters('" + $_.paramName + "')]"
#                    $index = $originalLine.LastIndexOf($_.paramValue)
#                    $ToPreserve = $originalLine.SubString(0, $index)
#                    $ToPreserveEnd = $originalLine.SubString($index + $_.paramValue.Length)
#                    #TODO - is it safe to always end the line with '",'?  shouldn't we use whatever the rest of the original line was after the quoted value?
#                    #$revisedLine = $ToPreserve + $dynamicParameter + '",'
#                    $revisedLine = $ToPreserve + $dynamicParameter + $ToPreserveEnd
#                }
#            }
#        }

        if ($originalLine -ne $revisedLine) {
            Write-Log $originalLine
            Write-Log $revisedLine
            Write-Log "***"
        }

        $UpdatedARMTemplate += $revisedLine
    }

    $UpdatedARMTemplate | Set-Content -Path $TemplatePath -Encoding UTF8
}

function ParameterizationPhase2 {
    Param (
        [Parameter(Mandatory = $true)][String] $TargetFolder,
        [Parameter(Mandatory = $true)][String] $TemplateFilename,
        [Parameter(Mandatory = $false)][object[]] $Parameters
    )

    <#
        .SYNOPSIS
        Parameterization Phase 2

        .DESCRIPTION
        Complete the remaining dynamic parameter replacements via JSON.

        .PARAMETER TargetFolder
        Target folder where the ARN template resides.

        .PARAMETER TemplateFilename
        ARM Template filename.

        .PARAMETER Parameters
        Array of parameter objects
    #>

    $TemplatePath = "$TargetFolder\$TemplateFilename"
    $armTemplate = Get-Content -Path $TemplatePath -Raw | ConvertFrom-Json

#    Function AddParameters( $armTemplate, $parameterName, $defaultValue ) {
#        # Build the defaultValue name and value pair dynamically
#        $parameterValue = "{""defaultValue"": ""$defaultValue"", ""type"": ""String""}"
#        $parameters = $parameterValue | ConvertFrom-Json
#        $armTemplate.parameters | Add-Member -NotePropertyName $parameterName -NotePropertyValue $parameters -Force
#    }

    <#
        In "resources" section, find "api" then "id" child node, and make it dynamic.
    
        For example, replace:
            [concat('/subscriptions/19538b17-fa74-44e8-8976-edb1c9a20ea2/providers/Microsoft.Web/locations/canadacentral/managedApis/', parameters('connections_sharepointonline_name'))]
        
        with:
            [concat('/subscriptions/', subscription().subscriptionId, '/providers/Microsoft.Web/locations/', resourceGroup().location, '/managedApis/', parameters('connections_sharepointonline_name'))]" 
    #>
    function UpdateIdField($armTemplate) {
        [String] $TokenAPI = "/managedApis"
        [String] $TokenParameter = "parameters("
        [String] $ToReplace = "[concat('/subscriptions/', subscription().subscriptionId, '/providers/Microsoft.Web/locations/', resourceGroup().location, '/managedApis/', "
    
        $armTemplate.Resources | ForEach-Object {
            if ($_.properties.api.id) {
                # make sure it is not null
                if ( $_.properties.api.id.Contains($TokenAPI) ) {
                    $index = $_.properties.api.id.LastIndexOf($TokenParameter)
                    $ToPreserve = $_.properties.api.id.SubString($index) # second part
                    $ReplaceIdValue = "$ToReplace$ToPreserve"
                    $_.properties.api.id = $ReplaceIdValue
                }
            }
        }
    }

    $armtemplate.parameters.psobject.Properties.Remove("certificates_AzureRunAsCertificate_base64Value")
    $trimmedResources = $armtemplate.resources | Where-Object {$_.type -ne "Microsoft.Automation/automationAccounts/jobs" `
                                                    -and $_.type -ne "Microsoft.Automation/automationAccounts/certificates" `
                                                    -and $_.type -ne "Microsoft.Automation/automationAccounts/connections" `
                                                    -and $_.type -ne "Microsoft.Automation/automationAccounts/connectionTypes" `
                                                    -and $_.type -ne "Microsoft.Automation/automationAccounts/modules" `
                                                    -and $_.type -ne "Microsoft.Automation/automationAccounts/runbooks" }
    $armtemplate.resources = $trimmedResources

#    if( $parameters ) {
#        $parameters | ForEach-Object {
#            AddParameters $armTemplate $_.paramName $_.paramValue    
#        }
#    }
    
    UpdateIdField $armTemplate
    
    # Save everything back to ARM Template in JSON format
    $armTemplate | ConvertTo-Json  -Depth 100  | Format-Json | ForEach-Object {
        [Regex]::Replace($_, 
            "\\u(?<Value>[a-zA-Z0-9]{4})", {
                param($m) ([char]([int]::Parse($m.Groups['Value'].Value,
                            [System.Globalization.NumberStyles]::HexNumber))).ToString() } ) } | Set-Content -Path $TemplatePath -Encoding UTF8
}

function TargetedParameterReplacements {
    Param (
        [Parameter(Mandatory = $true)][String] $TargetFolder,
        [Parameter(Mandatory = $true)][String] $TemplateFilename,
        [Parameter(Mandatory = $false)][object[]] $Parameters
    )

    <#
        .SYNOPSIS
        Parameterization of Automation Variables
        Parameterization of Logic App Params
        Parameterization of Connection display names

        .DESCRIPTION
        Make parameters for all AutomationVariable resources via JSON.

        .PARAMETER TargetFolder
        Target folder where the ARN template resides.

        .PARAMETER TemplateFilename
        ARM Template filename.

        .PARAMETER Parameters
        Array of parameters
    #>

    Function AddParameters( $armTemplate, $parameterName, $defaultValue ) {
        # Build the defaultValue name and value pair dynamically
        $parameterValue = "{""defaultValue"": ""$defaultValue"", ""type"": ""String""}"
        $parameters = $parameterValue | ConvertFrom-Json
        $armTemplate.parameters | Add-Member -NotePropertyName $parameterName -NotePropertyValue $parameters -Force
    }

    $TemplatePath = "$TargetFolder\$TemplateFilename"
    $armTemplate = Get-Content -Path $TemplatePath -Raw | ConvertFrom-Json

    # Create an ARM Template parameter for each parameter in the config.Parameters
    $Parameters | foreach-object {
        AddParameters $armTemplate $_.paramName $_.paramValue
    }

    # Find all the automationVariables in the template - if they appear in config.Parameters we need to set their value to refer to the ARM Template parameter
    $automationVariables = $armTemplate.resources | where-object {$_.type -eq "Microsoft.Automation/automationAccounts/variables"}
    $automationVariables | foreach-object {
        $avName = $_.name
        $temp = $avName.Substring($avName.LastIndexOf(", '/")+4)
        $name = $temp.Substring(0,$temp.LastIndexOf("'"))

        $nameMatch = $Parameters | where-object {$_.paramName -eq $name}
        if ($nameMatch) {
            $_.properties[0].value = "[parameters('$($name)')]"
        }

        $aliasMatch = $Parameters | where-object {$_.paramAliases -contains $name}
        if ($aliasMatch) {
            $_.properties[0].value = "[parameters('$($aliasMatch.paramName)')]"
        }
    }

    # Find all the Logic Apps in the template - if their parameters appear in config.Parameters we need to set their value to refer to the ARM Template parameter
    $logicApps = $armTemplate.resources | where-object {$_.type -eq "Microsoft.Logic/workflows"}
    $logicApps | foreach-object {
        $laParams = $_.properties.definition.parameters.psobject.properties
        $laParams | foreach-object {
            $name = $_.Name
            $nameMatch = $Parameters | where-object {$_.paramName -eq $name}
            if ($nameMatch) {
                #$_.properties[0].value = "[parameters('$($name)')]"
                $_.Value.psobject.Properties["defaultValue"].Value = "[parameters('$($name)')]"
            }

            $aliasMatch = $Parameters | where-object {$_.paramAliases -contains $name}
            if ($aliasMatch) {
                #$_.properties[0].value = "[parameters('$($name)')]"
                $_.Value.psobject.Properties["defaultValue"].Value = "[parameters('$($aliasMatch.paramName)')]"
            }
        }
    }

    # Find all the Connections in the template - if their displaynames appear in config.Parameters.paramValue we need to set their value to refer to the ARM Template parameter
    $connections = $armTemplate.resources | where-object {$_.type -eq "Microsoft.Web/connections"}
    $connections | foreach-object {
        $displayNameProperty = $_.properties.psobject.properties["displayName"]
        $valueMatch = $Parameters | where-object {$_.paramValue -eq $displayNameProperty.Value}
        if ($valueMatch) {
            $displayNameProperty.Value = "[concat('""',parameters('$($valueMatch.paramName)'),'""')]"
        }
    }

    # Save everything back to ARM Template in JSON format
    $armTemplate | ConvertTo-Json  -Depth 100  | Format-Json | ForEach-Object {
        [Regex]::Replace($_, 
            "\\u(?<Value>[a-zA-Z0-9]{4})", {
                param($m) ([char]([int]::Parse($m.Groups['Value'].Value,
                            [System.Globalization.NumberStyles]::HexNumber))).ToString() } ) } | Set-Content -Path $TemplatePath -Encoding UTF8
}

function Format-Json {
    [CmdletBinding(DefaultParameterSetName = 'Prettify')]
    Param(
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
        [String] $Json,

        [Parameter(ParameterSetName = 'Minify')]
        [switch] $Minify,

        [Parameter(ParameterSetName = 'Prettify')]
        [ValidateRange(1, 1024)]
        [int] $Indentation = 4,

        [Parameter(ParameterSetName = 'Prettify')]
        [switch] $AsArray
    )

    <#
        .SYNOPSIS
        Format the JSON text

        .DESCRIPTION
        Properly format the JSON text for output.

        .PARAMETER Json
        Mandatory: The JSON text to prettify.

        .PARAMETER Minify
        Optional: Returns the JSON string compressed.

        .PARAMETER Indentation
        Optional: The number of spaces (1..1024) to use for indentation. Defaults to 4.

        .PARAMETER AsArray
        Optional: If set, the output will be in the form of a string array, otherwise a single string is output.
    #>

    if ($PSCmdlet.ParameterSetName -eq 'Minify') {
        return ($Json | ConvertFrom-Json) | ConvertTo-Json -Depth 100 -Compress
    }

    # If the input JSON text has been created with ConvertTo-Json -Compress
    # then we first need to reconvert it without compression
    if ($Json -notmatch '\r?\n') {
        $Json = ($Json | ConvertFrom-Json) | ConvertTo-Json -Depth 100
    }

    $indent = 0
    $regexUnlessQuoted = '(?=([^"]*"[^"]*")*[^"]*$)'

    $result = $Json -split '\r?\n' |
    ForEach-Object {
        # If the line contains a ] or } character, 
        # we need to decrement the indentation level unless it is inside quotes.
        if ($_ -match "[}\]]$regexUnlessQuoted") {
            $indent = [Math]::Max($indent - $Indentation, 0)
        }

        # Replace all colon-space combinations by ": " unless it is inside quotes.
        $line = (' ' * $indent) + ($_.TrimStart() -replace ":\s+$regexUnlessQuoted", ': ')

        # If the line contains a [ or { character, 
        # we need to increment the indentation level unless it is inside quotes.
        if ($_ -match "[\{\[]$regexUnlessQuoted") {
            $indent += $Indentation
        }

        $line
    }

    if ($AsArray) { return $result }
    return $result -Join [Environment]::NewLine
}

function ParameterizeTemplate {
    Param (
        [Parameter(Mandatory = $true)][String] $TargetFolder,
        [Parameter(Mandatory = $true)][String] $TemplateFilename,
        [Parameter(Mandatory = $false)][object[]] $Parameters 
    )

    <#
        .SYNOPSIS
        Parameterize the ARM Template

        .DESCRIPTION
        Parameterize the ARM Template through several steps

        .PARAMETER TargetFolder
        Target folder where the ARN template resides.

        .PARAMETER TemplateFilename
        ARM Template filename.

        .PARAMETER ParameterDefaultValueFilename
        JSON file containing the Parameter Default Value Mapping.
    #>

    TargetedParameterReplacements -TargetFolder $TargetFolder -TemplateFilename $TemplateFilename -Parameters $Parameters
    ParameterizationPhase1 -TargetFolder $TargetFolder -TemplateFilename $TemplateFilename -Parameters $Parameters
    ParameterizationPhase2 -TargetFolder $TargetFolder -TemplateFilename $TemplateFilename -Parameters $Parameters
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

Write-Log "Start!"

$configFileInput = Read-Host "Enter the config.json full path (default - $configFile) "
if ($configFileInput -ne "") {
    $configFile = $configFileInput
}

ReadConfiguration -ConfigFile $configFile

# Test execution
ExportARMTemplateFromResourceGroup -SubscriptionId $config.SubscriptionId -ResourceGroupName $config.ResourceGroupName `
                                   -TargetFolder $config.TargetFolder -TemplateFilename $config.TemplateFilename
ParameterizeTemplate -TargetFolder $config.TargetFolder -TemplateFilename $config.TemplateFilename -Parameters $config.Parameters

Write-Log "Done!"