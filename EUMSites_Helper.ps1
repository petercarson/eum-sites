function LoadEnvironmentSettings()
{
    Param
    (
        [Parameter(Position=0,Mandatory=$false)][int] $environmentId
    )

    [xml]$config = Get-Content -Path "$DistributionFolder\sharepoint.config"

    if (-not($environmentId))
    {
        # Prompt for the environment defined in the config
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
    [string]$Global:ManagedCredentials = $environment.webApp.managedCredentials

    [string]$Global:EUMClientID = $environment.EUM.clientID
    [string]$Global:EUMSecret = $environment.EUM.secret
    [string]$Global:Domain_FK = $environment.EUM.domain_FK
    [string]$Global:SystemConfiguration_FK = $environment.EUM.systemConfiguration_FK
    [string]$Global:EUMURL = $environment.EUM.EUMURL

    Write-Host "Environment set to $($environment.name) - $($environment.webApp.adminSiteURL) `n" -ForegroundColor Cyan

    #Add required Client Dlls 
    Add-Type -Path "$DistributionFolder\Microsoft.SharePoint.Client.dll"
    Add-Type -Path "$DistributionFolder\Microsoft.SharePoint.Client.Taxonomy.dll"

	$Global:credentials = Get-StoredCredential -Target $managedCredentials 
    if ($credentials -eq $null) {
        $UserName = Read-Host "Enter the username to connect with"
        $Password = Read-Host "Enter the password for $UserName" -AsSecureString 
        $SaveCredentials = Read-Host "Save the credentials in Windows Credential Manager (Y/N)?"
        if (($SaveCredentials -eq "y") -or ($SaveCredentials -eq "Y")) {
            $temp = New-StoredCredential -Target $managedCredentials -UserName $UserName -SecurePassword $Password
        }
        $Global:SPCredentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($UserName, $Password)
    }
    else {
        $Global:SPCredentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($credentials.UserName, $credentials.Password)
    }
}

function ExportSiteColumns
{
    param ( [Parameter(Position=0,Mandatory=$true)][string]$ExportFile,
            [Parameter(Position=1,Mandatory=$true)][string]$GroupName,
	        [Parameter(Position=2,Mandatory=$true)][string]$SiteUrl)

    $TaxonomyFields = [System.Collections.ArrayList]@()

#    $web = Get-SPWeb -Identity $SiteUrl
    if ($SiteURL -ne $null -and $SiteURL -ne "" -and $SiteURL -ne $WebAppUrl) {
        $siteContext = New-Object Microsoft.SharePoint.Client.ClientContext($SiteURL)
	    $MixedModeAuthentication = $false;
        $siteContext.Credentials = $SPCredentials
    } else {
        $siteContext = $spContext
    }

    $web = $siteContext.Web
    $siteContext.Load($web.Fields)
    $siteContext.ExecuteQuery()

    $web.Fields | ForEach-Object {
        if (($_.Group -eq $GroupName) -and ($_.TypeAsString -eq "TaxonomyFieldType")) {
            $TaxonomyFields.Add($_.StaticName + "_0")
        }
    }

    #Create Export Files
    New-Item $ExportFile -type file -force

    #Export Site Columns to XML file
    Add-Content $ExportFile "<?xml version=`"1.0`" encoding=`"utf-8`"?>"
    Add-Content $ExportFile "`n<Fields>"

    $Web.Fields | ForEach-Object {
        if (($_.Group -eq $GroupName) -or ($TaxonomyFields -contains $_.Title)) {
            Add-Content $ExportFile $_.SchemaXml
        }
    }

    Add-Content $ExportFile "</Fields>"

    #$web.Dispose()
}

function ExportContentTypes
{
    param ( [Parameter(Position=0,Mandatory=$true)][string]$ExportFile,
            [Parameter(Position=1,Mandatory=$true)][string]$GroupName,
	        [Parameter(Position=2,Mandatory=$true)][string]$SiteUrl)

    #$web = Get-SPWeb -Identity $SiteUrl
    if ($SiteURL -ne $null -and $SiteURL -ne "" -and $SiteURL -ne $WebAppUrl) {
        $siteContext = New-Object Microsoft.SharePoint.Client.ClientContext($SiteURL)
	    $MixedModeAuthentication = $false;
        $siteContext.Credentials = $SPCredentials
    } else {
        $siteContext = $spContext
    }

    $web = $siteContext.Web
    $siteContext.Load($web.ContentTypes)
    $siteContext.ExecuteQuery()

    #Create Export Files
    New-Item $ExportFile -type file -force

    #Export Content Types to XML file
    Add-Content $ExportFile "<?xml version=`"1.0`" encoding=`"utf-8`"?>"
    Add-Content $ExportFile "`n<ContentTypes>"

    $web.ContentTypes | ForEach-Object {
        if ($_.Group -eq $GroupName) {
            Add-Content $ExportFile $_.SchemaXml
        }
    }

    Add-Content $ExportFile "</ContentTypes>"

    #$web.Dispose()
}


function ImportSiteColumns
{
    param ( [Parameter(Position=0,Mandatory=$true)][string]$ImportFile,
	        [Parameter(Position=1,Mandatory=$true)][string]$SiteUrl)

    if ($SiteURL -ne $null -and $SiteURL -ne "" -and $SiteURL -ne $WebAppUrl) {
        $siteContext = New-Object Microsoft.SharePoint.Client.ClientContext($SiteURL)
	    $MixedModeAuthentication = $false;
        $siteContext.Credentials = $SPCredentials
    } else {
        $siteContext = $spContext
    }

    $spWeb = $siteContext.Web
    $siteContext.Load($spWeb)

    #Get exported XML file
    $fieldsXML = [xml](Get-Content($ImportFile))

    #may need the termstore id if any fields are Managed Metadata
    $spTaxSession = [Microsoft.SharePoint.Client.Taxonomy.TaxonomySession]::GetTaxonomySession($siteContext)
    $spTaxSession.UpdateCache();
    $siteContext.Load($spTaxSession)
    $termStore = $spTaxSession.GetDefaultSiteCollectionTermStore()
    $siteContext.Load($termStore)

    try
    {
        $siteContext.ExecuteQuery()
    }
    catch
    {
        Write-host "Error while loading the Taxonomy Session " $_.Exception.Message -ForegroundColor Red 
        exit 1
    }


    $counter = 1

    $fieldsXML.Fields.Field | ForEach-Object {
    
        $isExistingField = $false
        $existingField = $spWeb.Fields.GetById($_.ID)
        $siteContext.Load($existingField)
        try 
        {
            $siteContext.ExecuteQuery()
            $isExistingField = $true
        }
        catch 
        {
            $isExistingField = $false
        }

        if ($isExistingField) {
            Write-Host "skipping pre-existing site column " $_.Name
        } else {
            #Configure core properties belonging to all column types
            $fieldXML = '<Field Type="' + $_.Type + '"
            Name="' + $_.Name + '"
            ID="' + $_.ID + '"
            Description="' + $_.Description + '"
            DisplayName="' + $_.DisplayName + '"
            StaticName="' + $_.StaticName + '"
            Group="' + $_.Group + '"
            Hidden="' + $_.Hidden + '"
            Required="' + $_.Required + '"
            Sealed="' + $_.Sealed + '"'
    
            #Configure optional properties belonging to specific column types – you may need to add some extra properties here if present in your XML file
            if ($_.ShowInDisplayForm) { $fieldXML = $fieldXML + "`n" + 'ShowInDisplayForm="' + $_.ShowInDisplayForm + '"'}
            if ($_.ShowInEditForm) { $fieldXML = $fieldXML + "`n" + 'ShowInEditForm="' + $_.ShowInEditForm + '"'}
            if ($_.ShowInListSettings) { $fieldXML = $fieldXML + "`n" + 'ShowInListSettings="' + $_.ShowInListSettings + '"'}
            if ($_.ShowInNewForm) { $fieldXML = $fieldXML + "`n" + 'ShowInNewForm="' + $_.ShowInNewForm + '"'}
        
            if ($_.EnforceUniqueValues) { $fieldXML = $fieldXML + "`n" + 'EnforceUniqueValues="' + $_.EnforceUniqueValues + '"'}
            if ($_.Indexed) { $fieldXML = $fieldXML + "`n" + 'Indexed="' + $_.Indexed + '"'}
            if ($_.Format) { $fieldXML = $fieldXML + "`n" + 'Format="' + $_.Format + '"'}
            if ($_.MaxLength) { $fieldXML = $fieldXML + "`n" + 'MaxLength="' + $_.MaxLength + '"' }
            if ($_.FillInChoice) { $fieldXML = $fieldXML + "`n" + 'FillInChoice="' + $_.FillInChoice + '"' }
            if ($_.NumLines) { $fieldXML = $fieldXML + "`n" + 'NumLines="' + $_.NumLines + '"' }
            if ($_.RichText) { $fieldXML = $fieldXML + "`n" + 'RichText="' + $_.RichText + '"' }
            if ($_.RichTextMode) { $fieldXML = $fieldXML + "`n" + 'RichTextMode="' + $_.RichTextMode + '"' }
            if ($_.IsolateStyles) { $fieldXML = $fieldXML + "`n" + 'IsolateStyles="' + $_.IsolateStyles + '"' }
            if ($_.AppendOnly) { $fieldXML = $fieldXML + "`n" + 'AppendOnly="' + $_.AppendOnly + '"' }
            if ($_.Sortable) { $fieldXML = $fieldXML + "`n" + 'Sortable="' + $_.Sortable + '"' }
            if ($_.RestrictedMode) { $fieldXML = $fieldXML + "`n" + 'RestrictedMode="' + $_.RestrictedMode + '"' }
            if ($_.UnlimitedLengthInDocumentLibrary) { $fieldXML = $fieldXML + "`n" + 'UnlimitedLengthInDocumentLibrary="' + $_.UnlimitedLengthInDocumentLibrary + '"' }
            if ($_.CanToggleHidden) { $fieldXML = $fieldXML + "`n" + 'CanToggleHidden="' + $_.CanToggleHidden + '"' }
            # commented out list since it seems to break metadata columns
		    #if ($_.List) { $fieldXML = $fieldXML + "`n" + 'List="' + $_.List + '"' }
            if ($_.ShowField) { $fieldXML = $fieldXML + "`n" + 'ShowField="' + $_.ShowField + '"' }
            if ($_.UserSelectionMode) { $fieldXML = $fieldXML + "`n" + 'UserSelectionMode="' + $_.UserSelectionMode + '"' }
            if ($_.UserSelectionScope) { $fieldXML = $fieldXML + "`n" + 'UserSelectionScope="' + $_.UserSelectionScope + '"' }
            if ($_.BaseType) { $fieldXML = $fieldXML + "`n" + 'BaseType="' + $_.BaseType + '"' }
            if ($_.Mult) { $fieldXML = $fieldXML + "`n" + 'Mult="' + $_.Mult + '"' }
            if ($_.ReadOnly) { $fieldXML = $fieldXML + "`n" + 'ReadOnly="' + $_.ReadOnly + '"' }
            if ($_.FieldRef) { $fieldXML = $fieldXML + "`n" + 'FieldRef="' + $_.FieldRef + '"' }    

            $fieldXML = $fieldXML + ">"
    
            #Create choices if choice column
            if ($_.Type -eq "Choice") {
                $fieldXML = $fieldXML + "`n<CHOICES>"
                $_.Choices.Choice | ForEach-Object {
                   $fieldXML = $fieldXML + "`n<CHOICE>" + $_ + "</CHOICE>"
                }
                $fieldXML = $fieldXML + "`n</CHOICES>"
            }
    
            #Set Default value, if specified  
            if ($_.Default) { $fieldXML = $fieldXML + "`n<Default>" + $_.Default + "</Default>" }
		
		    # Managed Metadata columns have this to specify which term set to use
		    if ($_.Customization) { 
            
                # Update the termstore ID in the InnerXML
                $termStoreID = $termStore.Id.ToString()

                $innerXML = $_.Customization.InnerXml
                $pattern = "<Property><Name>SspId</Name><Value xmlns:q1=`"http://www.w3.org/2001/XMLSchema`" p4:type=`"q1:string`" xmlns:p4=`"http://www.w3.org/2001/XMLSchema-instance`">.*</Value></Property><Property><Name>GroupId</Name>"
                $replacement = "<Property><Name>SspId</Name><Value xmlns:q1=`"http://www.w3.org/2001/XMLSchema`" p4:type=`"q1:string`" xmlns:p4=`"http://www.w3.org/2001/XMLSchema-instance`">" + $termStoreID +"</Value></Property><Property><Name>GroupId</Name>"

                $innerXML = $innerXML -replace $pattern, $replacement
                $fieldXML = $fieldXML + "`n<Customization>" + $innerXML + "</Customization>" 
            
            }
        
		    #End XML tag specified for this field
            $fieldXML = $fieldXML + "</Field>"
    
            #Create column on the site


            if (($_.Group -notlike  "Core") -and ($_.Group -ne "_Hidden")) {
                $addToDefaultView = $false
                $addFieldOptions = [Microsoft.SharePoint.Client.AddFieldOptions]::AddFieldInternalNameHint
                $newField = $spWeb.Fields.AddFieldAsXml($fieldXML.Replace("&","&amp;"), $addToDefaultView, $addFieldOptions)
                write-host $counter "Created site column" $_.DisplayName "on" $spWeb.Url
                $counter = $counter + 1
            }
    
        }
    }

    $siteContext.ExecuteQuery()
}

function ImportContentTypes
{
    param ( [Parameter(Position=0,Mandatory=$true)][string]$ImportFile,
	        [Parameter(Position=1,Mandatory=$true)][string] $SiteUrl)

    if ($SiteURL -ne $null -and $SiteURL -ne "" -and $SiteURL -ne $WebAppUrl) {
        $siteContext = New-Object Microsoft.SharePoint.Client.ClientContext($SiteURL)
	    $MixedModeAuthentication = $false;
        $siteContext.Credentials = $SPCredentials
    } else {
        $siteContext = $spContext
    }

    $spWeb = $siteContext.Web
    $siteContext.Load($spWeb)
    $siteContext.ExecuteQuery()

    $ctsXML = [xml](Get-Content($ImportFile))
    $ctsXML.ContentTypes.ContentType | ForEach-Object {
        $existingCT = $spWeb.ContentTypes.GetById($_.Id)
        $siteContext.Load($existingCT)
        $siteContext.ExecuteQuery()

        if ($existingCT.ServerObjectIsNull) {
            #Create Content Type object 
            write-host "Creating Content type" $_.Name
            $cti = New-Object Microsoft.SharePoint.Client.ContentTypeCreationInformation
            $cti.Name = $_.Name
            $cti.Id = $_.Id
            $cti.Description = $_.Description
            $cti.Group = $_.Group

            $newCT = $spWeb.ContentTypes.Add($cti)
            $siteContext.ExecuteQuery()
            $siteContext.Load($newCT)
            $newCTFieldLinks = $newCT.FieldLinks
            $siteContext.Load($newCTFieldLinks)
            $siteContext.ExecuteQuery()
    
            $_.Fields.Field  | ForEach-Object {
                $existingLink = $newCTFieldLinks.GetById($_.Id)
                $siteContext.Load($existingLink)
                $siteContext.ExecuteQuery()

                if($existingLink.ServerObjectIsNull)
                {
                    Write-Host "Create field link for " $_.Name " to content type " $newCT.Name -NoNewline
                    #Create a field link for the Content Type by getting an existing column
                    $fli = New-Object Microsoft.SharePoint.Client.FieldLinkCreationInformation
                    $field = $spWeb.Fields.GetById($_.Id)
                    $fli.Field = $field
                    $newFL = $newCT.FieldLinks.Add($fli)
                    #Check to see if column should be Optional, Required or Hidden
                    if ($_.Required -eq "TRUE") {$newFL.Required = $true}
                    if ($_.Hidden -eq "TRUE") {$newFL.Hidden = $true}
                    $newCT.Update($true)
                    try 
                    {
                        $siteContext.ExecuteQuery()
                        Write-Host "... Success" -ForegroundColor Green
                    }
                    catch [Exception]
                    {
					    Write-Host "... ERROR " -ForegroundColor Red 
                        Write-Host $_.Exception.Message -ForegroundColor Red 
                        if ($_.Exception.ServerErrorTraceCorrelationId -ne $null -and $_.Exception.ServerErrorTraceCorrelationId -ne '') {
                            #Write-Host "Server Correlation Id " $_.Exception.ServerErrorTraceCorrelationId -ForegroundColor Red 
                        }
                    }
                } #if $existingLink.ServerObjectIsNull
            } #foreach field
    
            write-host "Content type" $newCT.Name "has been created"
        } #if $existingCT.ServerObjectIsNull
        else 
        {
            write-host "Skipping existing Content type" $_.Name
        }
    } #foreach ContentType

}

