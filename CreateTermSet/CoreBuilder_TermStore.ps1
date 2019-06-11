# ----------------------------------------------------------
# 
# Copyright Envision IT Inc. https://www.envisionit.com
# Licensed under a Creative Commons Attribution-ShareAlike 3.0 Unported License
# https://creativecommons.org/licenses/by-sa/3.0/deed.en_US
# 
# ----------------------------------------------------------

function ImportTermStore
{
    Param([Parameter(Position=0,Mandatory=$true)][string] $ImportFile,
        [Parameter(Position=1,Mandatory=$true)][string] $TermStoreGroup,
        [Parameter(Position=2,Mandatory=$false)][string] $SiteURL)

	$termStore = Get-TermStoreInfo $spContext
	$termsXML = Get-TermsToImport $ImportFile

	Create-Groups $spContext $termStore $termsXML
}

function SetTermsStoreNavigation
{
    Param([Parameter(Position=0,Mandatory=$true)][string] $TermStoreGroup,
        [Parameter(Position=1,Mandatory=$true)][string] $TermSetName,
        [Parameter(Position=2,Mandatory=$false)][string] $SiteURL)

	# -----------------------------------------------------------------
	# Set Site Navigation Available for Tagging
	# -----------------------------------------------------------------

	$siteNavTermSet = "Site Navigation"

    $spTaxSession = [Microsoft.SharePoint.Client.Taxonomy.TaxonomySession]::GetTaxonomySession($spContext)
    $spTaxSession.UpdateCache();
    $spContext.Load($spTaxSession)
    $termStore = $spTaxSession.GetDefaultSiteCollectionTermStore()
    $spContext.Load($termStore)

    try
    {
        $spContext.ExecuteQuery()
    }
    catch
    {
        Write-host "Error while loading the Taxonomy Session " $_.Exception.Message -ForegroundColor Red 
        exit 1
    }

	#$termStore = $spTaxSession.TermStores["Managed Metadata Service"]
    $groups = $termStore.Groups
    $spContext.Load($groups)
    $spContext.ExecuteQuery()
    $group = $groups | Where-Object {$_.Name -eq $groupName}
    $termSetsForCheck = $group.TermSets
    $spContext.Load($group.TermSets)
    $spContext.ExecuteQuery()
	$termSet = $group.TermSets | Where-Object {$_.Name -eq $siteNavTermSet}

	$termSet.IsAvailableForTagging = $true
	$termStore.CommitAll()
    $spContext.ExecuteQuery()

	# -----------------------------------------------------------------
	# Assign it to the site
	# -----------------------------------------------------------------



    $navSettings = New-Object Microsoft.SharePoint.Client.Publishing.Navigation.WebNavigationSettings -ArgumentList $spContext, $spContext.Web

    #Global Navigation
    $navSettings.GlobalNavigation.Source = 2
    $navSettings.GlobalNavigation.TermStoreId = $termStore.Id
    $navSettings.GlobalNavigation.TermSetId = $termSet.Id

    $navSettings.Update($spTaxSession)

    $spContext.ExecuteQuery()
    Write-Host "Set site navigation to $TermStoreGroup"
}


function LinkSiteColumnToTaxonomy
{
  Param([Parameter(Position=0,Mandatory=$true)][string] $SiteColumn,
        [Parameter(Position=1,Mandatory=$true)][string] $TermSetId,
        [Parameter(Position=2,Mandatory=$false)][switch] $AllowMultipleValues)

    Write-Host "Linking site column " $SiteColumn " to term set with ID " $TermSetId
    $session = Get-SPTaxonomySession -site $WebAppURL
    $spWeb = Get-SPWeb $WebAppURL
    $field = $spWeb.Fields[$SiteColumn]
    If ($AllowMultipleValues.IsPresent)
    {
        $field.AllowMultipleValues = $true
    }
    $field.SspId =  $session.TermStores["Managed Metadata Service"].Id
    $field.TermSetId = $TermSetId
    $field.Update()

}

function ExportTermStore
{
    Param([Parameter(Position=0,Mandatory=$true)][string] $PathToExportXMLTerms,
        [Parameter(Position=1,Mandatory=$true)][string] $XMLTermsFileName,
        [Parameter(Position=2,Mandatory=$false)][string] $GroupToExport,
        [Parameter(Position=3,Mandatory=$false)][string] $SiteURL,
        [switch] $ExcludeIDs)

	$termStore = Get-TermStoreInfo $spContext
	$xmlFile = Get-XMLTermStoreTemplateToFile $termStore.Name $PathToExportXMLTerms
	Get-XMLFileObjectTemplates $xmlFile
    ExportTaxonomy $spContext $termStore $xmlFile $GroupToExport $PathToExportXMLTerms $XMLTermsFileName -ExcludeIDs:$ExcludeIDs
}

<#
 .NOTES
    Created By Paul Matthews, with original input from Luis Manez and Kevin Beckett.

 .LINK 
    http://cannonfodder.wordpress.com -Paul Matthews Blog Post About this.
    http://geeks.ms/blogs/lmanez -Luis Manez Blog

 .SYNOPSIS
    Imports an exported Taxonomy XML file to SharePoint On-Prem or 365 environment.
 
 .DESCRIPTION
    The Import-Taxonomy.ps1 function will read through a given XML File and import Groups, TermSets, Terms 
    into the SharePoint Term Store if they do not exist. Works for Online and On-Prem environments

 .PARAMETER AdminUser
    The user who has adminitrative access to the term store. (e.g On-Prem: Domain\user 365:user@sp.com)

 .PARAMETER AdminPassword
    The password for the Admin User.

 .PARAMETER AdminUrl
    The URL of Central Admin for On-Prem or Admin site for 365

 .PARAMETER FilePathOfExportXMLTerms
    The path you wish to save the XML Output to. This path must exist.

 .PARAMETER PathToSPClientdlls
   The script requires to call the following dlls:
   Microsoft.SharePoint.Client.dll
   Microsoft.SharePoint.Client.Runtime.dll
   Microsoft.SharePoint.Client.Taxonomy.dll

   E.g C:\Program Files\Common Files\microsoft shared\Web Server Extensions\15\ISAPI

 .EXAMPLE
    This imports the XML into the SharePoint term store.
    ./Import-Taxonomy.ps1 -AdminUser user@sp.com -AdminPassword password -AdminUrl https://sp-admin.onmicrosoft.com -FilePathOfExportXMLTerms c:\myTerms\exportedterms.xml -PathToSPClientdlls "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\15\ISAPI"

#>

function Get-TermStoreInfo($spContext)
{
	 $spTaxSession = [Microsoft.SharePoint.Client.Taxonomy.TaxonomySession]::GetTaxonomySession($spContext)
	 $spTaxSession.UpdateCache();
	 $spContext.Load($spTaxSession)

	 try
	 {
	 $spContext.ExecuteQuery()
	 }
	 catch
	 {
	  Write-host "Error while loading the Taxonomy Session " $_.Exception.Message -ForegroundColor Red 
	  exit 1
	 }

	 if($spTaxSession.TermStores.Count -eq 0){
	  write-host "The Taxonomy Service is offline or missing" -ForegroundColor Red
	  exit 1
	 }

	 $termStores = $spTaxSession.TermStores
	 $spContext.Load($termStores)

	 try
	 {
	  $spContext.ExecuteQuery()
	  $termStore = $termStores[0]
	  $spcontext.Load($termStore)
	  Write-Host "Connected to TermStore: $($termStore.Name) ID: $($termStore.Id)"
	 }
	 catch
	 {
	  Write-host "Error details while getting term store ID" $_.Exception.Message -ForegroundColor Red
	  exit 1
	 }

	 return $termStore
}

function Get-TermsToImport($xmlTermsPath)
{
	 [Reflection.Assembly]::LoadWithPartialName("System.Xml.Linq") | Out-Null

	 try
	 {
		 $xDoc = [System.Xml.Linq.XDocument]::Load($xmlTermsPath, [System.Xml.Linq.LoadOptions]::None)
		 return $xDoc
	 }
	 catch
	 {
		  Write-Host "Unable to read ExportedTermsXML. Exception:$_.Exception.Message" -ForegroundColor Red
		  exit 1
	 }
}

function Delete-Group
{
    Param([Parameter(Position=0,Mandatory=$true)][Microsoft.SharePoint.Client.ClientContext] $spContext,
        [Parameter(Position=1,Mandatory=$true)][Microsoft.SharePoint.Client.Taxonomy.TermGroup] $group)

    Write-Host "    Deleting taxonomy group " $group.Name
    $termSets = $group.TermSets
    $spContext.Load($termSets)
    $spContext.ExecuteQuery()
    $termSets | foreach-object { Delete-Termset $spContext $_ }
    $group.DeleteObject()
    $spContext.ExecuteQuery()
}

function Delete-Termset
{
    Param([Parameter(Position=0,Mandatory=$true)][Microsoft.SharePoint.Client.ClientContext] $spContext,
        [Parameter(Position=1,Mandatory=$true)][Microsoft.SharePoint.Client.Taxonomy.TermSet] $termSet)

    Write-Host "        Deleting taxonomy termset " $termSet.Name
    $terms = $termSet.Terms
    $spContext.Load($terms)
    $spContext.ExecuteQuery()
    $terms | foreach-object { Delete-Term $spContext $_ }
    $termSet.DeleteObject()
    $spContext.ExecuteQuery()
}

function Delete-Term
{
    Param([Parameter(Position=0,Mandatory=$true)][Microsoft.SharePoint.Client.ClientContext] $spContext,
        [Parameter(Position=1,Mandatory=$true)][Microsoft.SharePoint.Client.Taxonomy.Term] $term)

    Write-Host "            Deleting taxonomy term " $term.Name
    $term.DeleteObject()
    $spContext.ExecuteQuery()
}

function Create-Groups($spContext, $termStore, $termsXML){
     foreach($groupNode in $termsXML.Descendants("Group"))
     {
        if ($groupNode.Attribute("IsSiteCollectionGroup").Value -eq $true)
        {
            $site = $spContext.get_site()
            $group = $termStore.GetSiteCollectionGroup($site, $true)
            $spContext.Load($group)
            $spContext.ExecuteQuery()
        }
        else
        {
            $name = $groupNode.Attribute("Name").Value
            $description = $groupNode.Attribute("Description").Value;
            $groupId = $groupNode.Attribute("Id").Value;
            if (($groupId -eq $null) -or ($groupId -eq ""))
            {
                $groupGuid = [guid]::NewGuid()
            }
            else
            {
                $groupGuid = [System.Guid]::Parse($groupId);
            }

            Write-Host "Processing Group: $name ID: $groupId ..." -NoNewline

            $group = $termStore.GetGroup($groupGuid);
            $spContext.Load($group);
        
            try
            {
                $spContext.ExecuteQuery()
            }
            catch
            {
                Write-host "Error while finding if " $name " group already exists. " $_.Exception.Message -ForegroundColor Red 
                exit 1
            }


            if (!$group.ServerObjectIsNull) { 
                #group with give $guid already exists so delete it and recreate it
                Delete-Group $spContext $group
            } 
            else
            {
                #group with given $guid does not exist - but is there already one with the same name that we need to replace?
                #  this can happen with sharepoint online because People group is automatically created
                #  when a site collection is created
                $groupsForCheck = $termStore.Groups
                $spcontext.Load($groupsForCheck)
                $spcontext.ExecuteQuery()
                $groupByName = $groupsForCheck | Where-Object {$_.Name -eq $name}
                if ($groupByName -ne $null) {
                    #$groupByName.DeleteObject()
                    #$spcontext.ExecuteQuery()
                    Delete-Group $spContext $groupByName
                    Write-host "Deleted existing group named " $name " so that it can be recreated with correct GUID"
                }
            }

            $group = $termStore.CreateGroup($name, $groupGuid);
            $spContext.Load($group);
            try
            {
                $spContext.ExecuteQuery();
		        write-host "Inserted" -ForegroundColor Green
            }
            catch
            {
                Write-host "Error creating new Group " $name " " $_.Exception.Message -ForegroundColor Red 
                exit 1
            }
        }
	
	    Create-TermSets $termsXML $group $termStore $spContext

     }

     try
     {
         $termStore.CommitAll();
         $spContext.ExecuteQuery();
     }
     catch
     {
       Write-Host "Error commiting changes to server. Exception:$_.Exception.Message" -foregroundcolor red
       exit 1
     }
}

function Create-TermSets($termsXML, $group, $termStore, $spContext) {
	
    $termSets = $termsXML.Descendants("TermSet") | Where { (($_.Parent.Parent.Attribute("Name").Value -eq $group.Name) -or ($_.Parent.Parent.Attribute("IsSiteCollectionGroup").Value)) }

	foreach ($termSetNode in $termSets)
    {
        $errorOccurred = $false

		$name = $termSetNode.Attribute("Name").Value;
        $id = $termSetNode.Attribute("Id").Value
        if (($id -eq $null) -or ($id -eq ""))
        {
            $guid = [guid]::NewGuid()
        }
        else
        {
            $guid = [System.Guid]::Parse($id);
        }
        $description = $termSetNode.Attribute("Description").Value;
        $customSortOrder = $termSetNode.Attribute("CustomSortOrder").Value;
        Write-host "Processing TermSet $name ... " -NoNewLine
		
		$termSet = $termStore.GetTermSet($guid);
        $spcontext.Load($termSet);
                
        try
        {
            $spContext.ExecuteQuery();
        }
        catch
        {
            Write-host "Error while finding if " $name " termset already exists. " $_.Exception.Message -ForegroundColor Red 
            exit 1
        }
		
		if ($termSet.ServerObjectIsNull) 
        {
            #termset with given $guid does not exist - but is there already one with the same name that we need to replace?
            #  this can happen with sharepoint online because Site Navigation and Wiki Categories termsets are automatically created
            #  when a site collection is created
            $termSetsForCheck = $group.TermSets
            $spcontext.Load($termSetsForCheck)
            $spcontext.ExecuteQuery()
            $termSetByName = $termSetsForCheck | Where-Object {$_.Name -eq $name}
            if ($termSetByName -ne $null) {
                $termSetByName.DeleteObject()
                $spcontext.ExecuteQuery()
                Write-host "Deleted existing termset named " $name " so that it can be recreated with correct GUID"
            }

			$termSet = $group.CreateTermSet($name, $guid, $termStore.DefaultLanguage);
            $termSet.Description = $description;
            
            if($customSortOrder -ne $null)
            {
                $termSet.CustomSortOrder = $customSortOrder
            }
            
            $termSet.IsAvailableForTagging = [bool]::Parse($termSetNode.Attribute("IsAvailableForTagging").Value);
            $termSet.IsOpenForTermCreation = [bool]::Parse($termSetNode.Attribute("IsOpenForTermCreation").Value);

            if($termSetNode.Element("CustomProperties") -ne $null)
            {
                foreach($custProp in $termSetNode.Element("CustomProperties").Elements("CustomProperty"))
                {
                    $termSet.SetCustomProperty($custProp.Attribute("Key").Value, $custProp.Attribute("Value").Value)
                }
            }
            
            try
            {
                $spContext.ExecuteQuery();
            }
            catch
            {
                Write-host "Error occured while create Term Set" $name $_.Exception.Message -ForegroundColor Red
                $errorOccurred = $true
            }

            write-host "created" -ForegroundColor Green
		}
		else {
			write-host "Already exists" -ForegroundColor Yellow
		}
			
        
        if(!$errorOccurred)
        {
            if ($termSetNode.Element("Terms") -ne $null) 
            {
               foreach ($termNode in $termSetNode.Element("Terms").Elements("Term"))
               {
                  Create-Term $termNode $null $termSet $termStore $termStore.DefaultLanguage $spContext
               }
            }	
        }						
    }
}


function Create-Term($termNode, $parentTerm, $termSet, $store, $lcid, $spContext){
    $id = $termNode.Attribute("Id").Value
    if (($id -eq $null) -or ($id -eq ""))
    {
        $guid = [guid]::NewGuid()
    }
    else
    {
        $guid = [System.Guid]::Parse($id);
    }
    $name = $termNode.Attribute("Name").Value;
    $term = $termSet.GetTerm($guid);
    $errorOccurred = $false
	
   
    $spContext.Load($term);


    try
    {
        $spContext.ExecuteQuery();
    }
    catch
    {
        Write-host "Error while finding if " $name " term id already exists. " $_.Exception.Message -ForegroundColor Red 
        exit 1
    }

     write-host "Processing Term $name ..." -NoNewLine 
    if($term.ServerObjectIsNull)
    {
	    if ($parentTerm -ne $null) 
        {
            $term = $parentTerm.CreateTerm($name, $lcid, $guid);
        }
        else 
        {
        
            $term = $termSet.CreateTerm($name, $lcid, $guid);
        }


        $customSortOrder = $termNode.Attribute("CustomSortOrder").Value;
        $description = $termNode.Element("Descriptions").Element("Description").Attribute("Value").Value;
        $term.SetDescription($description, $lcid);
        $term.IsAvailableForTagging = [bool]::Parse($termNode.Attribute("IsAvailableForTagging").Value);
    
        if($customSortOrder -ne $null)
        {
            $term.CustomSortOrder = $customSortOrder
        }


        if($termNode.Element("CustomProperties") -ne $null)
        {
            foreach($custProp in $termNode.Element("CustomProperties").Elements("CustomProperty"))
            {
                $term.SetCustomProperty($custProp.Attribute("Key").Value, $custProp.Attribute("Value").Value)
            }
        }

        if($termNode.Element("LocalCustomProperties") -ne $null)
        {
            foreach($localCustProp in $termNode.Element("LocalCustomProperties").Elements("LocalCustomProperty"))
            {
                $term.SetLocalCustomProperty($localCustProp.Attribute("Key").Value, $localCustProp.Attribute("Value").Value)
            }
        }

        try
        {
            $spContext.Load($term);
            $spContext.ExecuteQuery();
	        write-host " created" -ForegroundColor Green	
	    }
        catch
        {
            Write-host "Error occured while create Term" $name $_.Exception.Message -ForegroundColor Red
            $errorOccurred = $true
        }
    }
    else
    {
     write-host "Already exists" -ForegroundColor Yellow
#     $pTermSet = $term.TermSet
#     $spContext.Load($pTermSet)
#     $spContext.ExecuteQuery()
#     $pTermSet.Name
    }
     
    if(!$errorOccurred)
    {
	    if ($termNode.Element("Terms") -ne $null) 
        {
            foreach ($childTermNode in $termNode.Element("Terms").Elements("Term")) 
            {
                Create-Term $childTermNode $term $termSet $store $lcid $spContext
            }
        }

    }
}


<#
 .NOTES
    Created By Paul Matthews, with original input from Luis Manez and Kevin Beckett.

 .LINK 
    http://cannonfodder.wordpress.com -Paul Matthews Blog Post About this.
    http://geeks.ms/blogs/lmanez -Luis Manez Blog

 .SYNOPSIS
    Exports a Taxonomy Group, or Groups from SharePoint On-Prem or 365 environment and saves to XML File.
 
 .DESCRIPTION
    The Export-Taxonomy.ps1 function will read through a given SharePoint Term Store Taxonomy, or given 
    Term Store Group and export the information to XML File.

 .PARAMETER AdminUrl
    The URL of Central Admin for On-Prem or Admin site for 365

 .PARAMETER PathToExportXMLTerms
    The path you wish to save the XML Output to. This path must exist.

 .PARAMETER XMLTermsFileName
   The name of the XML file to save. If the file already exists then it will be overwritten.

 .PARAMETER PathToSPClientdlls
   The script requires to call the following dlls:
   Microsoft.SharePoint.Client.dll
   Microsoft.SharePoint.Client.Runtime.dll
   Microsoft.SharePoint.Client.Taxonomy.dll

   (e.g., C:\Program Files\Common Files\microsoft shared\Web Server Extensions\15\ISAPI)

 .PARAMETER GroupToExport
  An optional parameter, if included only the Group will be exported. If omitted then the entire termstore will be written to XML.

 .EXAMPLE
    This exports the entire termstore.
    ./Export-Taxonomy.ps1 -AdminUser user@sp.com -AdminPassword password -AdminUrl https://sp-admin.onmicrosoft.com -PathToExportXMLTerms c:\myTerms -XMLTermsFileName exportterms.xml -PathToSPClientdlls "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\15\ISAPI"

 .EXAMPLE
    This exports just the Term Store Group 'Client Group Terms'
    ./Export-Taxonomy.ps1 -AdminUser user@sp.com -AdminPassword password -AdminUrl https://sp-admin.onmicrosoft.com -PathToExportXMLTerms c:\myTerms -XMLTermsFileName exportterms.xml -PathToSPClientdlls "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\15\ISAPI" -GroupToExport 'Client Group Terms'
 
#>

  function Get-XMLTermStoreTemplateToFile($termStoreName, $path){
 ## Set up an xml template used for creating your exported xml
    $xmlTemplate = '<TermStores>
    	<TermStore Name="' + $termStoreName + '" IsOnline="True" WorkingLanguage="1033" DefaultLanguage="1033" SystemGroup="c6fb3e37-0997-42b1-8e3c-2706a36adbc4">
    		<Groups>
				<Group Id="" Name="" Description="" IsSystemGroup="False" IsSiteCollectionGroup="False">
	    			<TermSets>
						<TermSet Id="" Name="" Description="" Contact="" IsAvailableForTagging="" IsOpenForTermCreation="" CustomSortOrder="False">
                            <CustomProperties>
                                <CustomProperty Key="" Value=""/>
                            </CustomProperties>
		    				<Terms>
								<Term Id="" Name="" IsDeprecated="" IsAvailableForTagging="" IsKeyword="" IsReused="" IsRoot="" IsSourceTerm="" CustomSortOrder="False">
                                    <Descriptions>
                                      <Description Language="1033" Value="" />
                                    </Descriptions>
                                    <CustomProperties>
                                        <CustomProperty Key="" Value="" />
                                    </CustomProperties>
                                    <LocalCustomProperties>
                                        <LocalCustomProperty Key="" Value="" />
                                    </LocalCustomProperties>
                                    <Labels>
                                      <Label Value="" Language="1033" IsDefaultForLanguage="" />
                                    </Labels>
                                    <Terms>                                      
                                       <Term Id="" Name="" IsDeprecated="" IsAvailableForTagging="" IsKeyword="" IsReused="" IsRoot="" IsSourceTerm="" CustomSortOrder="False">
                                            <Descriptions>
                                              <Description Language="1033" Value="" />
                                            </Descriptions>
                                            <CustomProperties>
                                                <CustomProperty Key="" Value="" />
                                            </CustomProperties>
                                            <LocalCustomProperties>
                                                <LocalCustomProperty Key="" Value="" />
                                            </LocalCustomProperties>
                                            <Labels>
                                              <Label Value="" Language="1033" IsDefaultForLanguage="" />
                                            </Labels>
                                       </Term>
                                    </Terms>
                                </Term>
							</Terms>							
		    			</TermSet>
					</TermSets>
	    		</Group>
    		</Groups>	
    	</TermStore>
    </TermStores>' 

try
{
	 #Save Template to disk
	 $xmlTemplate | Out-File($path + "\Template.xml")
 
	 #Load file and return
	 $xml = New-Object XML
	 $xml.Load($path + "\Template.xml")
	 return $xml
	 }
	 catch{
	  Write-host "Error creating Template file. " $_.Exception.Message -ForegroundColor Red
	  exit 1
	 }
 
}

function Get-XMLFileObjectTemplates($xml){
    #Grab template elements so that we can easily copy them later.
    $global:xmlGroupT = $xml.selectSingleNode('//Group[@Id=""]')  
    $global:xmlTermSetT = $xml.selectSingleNode('//TermSet[@Id=""]')  
    $global:xmlTermT = $xml.selectSingleNode('//Term[@Id=""]')
    $global:xmlTermLabelT = $xml.selectSingleNode('//Label[@Value=""]')
    $global:xmlTermDescriptionT = $xml.selectSingleNode('//Description[@Value=""]')
    $global:xmlTermCustomPropertiesT = $xml.selectSingleNode('//CustomProperty[@Key=""]')
    $global:xmlTermLocalCustomPropertiesT = $xml.selectSingleNode('//LocalCustomProperty[@Key=""]')
}

function Get-TermByGuid($xml, $guid, $parentTermsetGuid) {
    if ($parentTermsetGuid) {
        return  $xml.selectnodes('//Term[@Id="' + $guid + '"]')
    } else {
        return  $xml.selectnodes('//TermSet[@Id="' + $guid + '"]') 
    }
}

function Clean-Template($xml) {
    #Do not cleanup empty description nodes (this is the default state)

    ## Empty Term.Labels.Label
    $xml.selectnodes('//Label[@Value=""]') | ForEach-Object {
        $parent = $_.get_ParentNode()
        $parent.RemoveChild($_)  | Out-Null      
    } 
    ## Empty Term
    $xml.selectnodes('//Term[@Name=""]') | ForEach-Object {
        $parent = $_.get_ParentNode()
        $parent.RemoveChild($_)  | Out-Null      
    } 
    ## Empty TermSet
    $xml.selectnodes('//TermSet[@Name=""]') | ForEach-Object {
        $parent = $_.get_ParentNode()
        $parent.RemoveChild($_)  | Out-Null      
    } 
    ## Empty Group
    $xml.selectnodes('//Group[@Name=""]') | ForEach-Object {
        $parent = $_.get_ParentNode()
        $parent.RemoveChild($_)   | Out-Null     
    }
    ## Empty Custom Properties
    $xml.selectnodes('//CustomProperty[@Key=""]') | ForEach-Object {
     $parent = $_.get_ParentNode()
     $parent.RemoveChild($_) | Out-Null
    }

    ## Empty Local Custom proeprties
    $xml.selectnodes('//LocalCustomProperty[@Key=""]') | ForEach-Object {
    $parent = $_.get_ParentNode()
     $parent.RemoveChild($_) | Out-Null
    }

    $xml.selectnodes('//Descriptions')| ForEach-Object {
     $childNodes = $_.ChildNodes.Count
     if($childNodes -gt 1)
     {
        $_.RemoveChild($_.ChildNodes[0]) | Out-Null
     }
    }

    While ($xml.selectnodes('//Term[@Name=""]').Count -gt 0)
    {
        #Cleanup the XML, remove empty Term Nodes
        $xml.selectnodes('//Term[@Name=""]').RemoveAll() | Out-Null
    }   
}

function Get-TermSets($spContext, $xmlnewGroup, $termSets, $xml, [switch] $ExcludeIDs){
 
 $termSets | ForEach-Object{
    #Add each termset to the export xml
    $xmlNewSet = $global:xmlTermSetT.Clone()
    #Replace SharePoint ampersand with regular
    $xmlNewSet.Name = $_.Name.replace("＆", "&")
   
    if (-Not $ExcludeIDs.IsPresent)
    {
        $xmlNewSet.Id = $_.Id.ToString()
    }
   
    if ($_.CustomSortOrder -ne $null) 
    { 
        $xmlNewSet.CustomSortOrder = $_.CustomSortOrder.ToString()            
    }


    foreach($customprop in $_.CustomProperties.GetEnumerator())
    {
        ## Clone Term customProp node
        $xmlNewTermCustomProp = $global:xmlTermCustomPropertiesT.Clone()    

        $xmlNewTermCustomProp.Key = $($customProp.Key)
        $xmlNewTermCustomProp.Value = $($customProp.Value)
        $xmlNewSet.CustomProperties.AppendChild($xmlNewTermCustomProp) | Out-Null 
    }

    $xmlNewSet.Description = $_.Description.ToString()
    $xmlNewSet.Contact = $_.Contact.ToString()
    $xmlNewSet.IsOpenForTermCreation = $_.IsOpenForTermCreation.ToString()  
    $xmlNewSet.IsAvailableForTagging = $_.IsAvailableForTagging.ToString()  
    $xmlNewGroup.TermSets.AppendChild($xmlNewSet) | Out-Null

    Write-Host "Adding TermSet " -NoNewline
    Write-Host $_.name -ForegroundColor Green -NoNewline
    Write-Host " to Group " -NoNewline
    Write-Host $xmlNewGroup.Name -ForegroundColor Green

    $spContext.Load($_.Terms)
    try
    {
     $spContext.ExecuteQuery()
    }
    catch
    {
     Write-host "Error while loading Terms for TermSet " $_.name " " $_.Exception.Message -ForegroundColor Red
     exit 1
    }
    # Recursively loop through all the terms in this termset
    Get-Terms $spContext $xmlNewSet $_.Terms $xml -ExcludeIDs:$ExcludeIDs
 }

}

function Get-Terms($spContext, $parent, $terms, $xml, [switch] $ExcludeIDs){
 #Terms could be either the original termset or parent term with children terms
 $terms | ForEach-Object{
    #Create a new term xml Element
    $xmlNewTerm = $global:xmlTermT.Clone()
    #Replace SharePoint ampersand with regular
    $xmlNewTerm.Name = $_.Name.replace("＆", "&")
    if (-Not $ExcludeIDs.IsPresent)
    {
        $xmlNewTerm.id = $_.Id.ToString()
    }
    $xmlNewTerm.IsAvailableForTagging = $_.IsAvailableForTagging.ToString()
    $xmlNewTerm.IsKeyword = $_.IsKeyword.ToString()
	$xmlNewTerm.IsReused = $_.IsReused.ToString()
	$xmlNewTerm.IsRoot = $_.IsRoot.ToString()
    $xmlNewTerm.IsSourceTerm = $_.IsSourceterm.ToString()
    $xmlNewTerm.IsDeprecated = $_.IsDeprecated.ToString()

    if($_.CustomSortOrder -ne $null)
    {
        $xmlNewTerm.CustomSortOrder = $_.CustomSortOrder.ToString()  
    }

    #Custom Properties
    foreach($customprop in $_.CustomProperties.GetEnumerator())
    {
        # Clone Term customProp node
        $xmlNewTermCustomProp = $global:xmlTermCustomPropertiesT.Clone()    
        
        $xmlNewTermCustomProp.Key = $($customProp.Key)
        $xmlNewTermCustomProp.Value = $($customProp.Value)
        $xmlNewTerm.CustomProperties.AppendChild($xmlNewTermCustomProp)  | Out-Null
    }

    #Local Properties
    foreach($localProp in $_.LocalCustomProperties.GetEnumerator())
    {
       # Clone Term LocalProp node
       $xmlNewTermLocalCustomProp = $global:xmlTermLocalCustomPropertiesT.Clone()    

       $xmlNewTermLocalCustomProp.Key = $($localProp.Key)
       $xmlNewTermLocalCustomProp.Value = $($localProp.Value)
       $xmlNewTerm.LocalCustomProperties.AppendChild($xmlNewTermLocalCustomProp) | Out-Null
    }

    if($_.Description -ne ""){
        $xmlNewTermDescription = $global:xmlTermDescriptionT.Clone()    
        $xmlNewTermDescription.Value = $_.Description
        $xmlNewTerm.Descriptions.AppendChild($xmlNewTermDescription) |Out-Null
    }
    
    $spContext.Load($_.Labels)
    $spContext.Load($_.TermSet)
    $spContext.Load($_.Parent)
    $spContext.Load($_.Terms)

    try
    {
      $spContext.ExecuteQuery()
    }
    catch
    {
      Write-host "Error while loaded addition information for Term " $xmlNewTerm.Name "  " $_.Exception.Message -ForegroundColor Red
      exit 1
    }

    foreach($label in $_.Labels)
     {  
        ## Clone Term Label node
        $xmlNewTermLabel = $global:xmlTermLabelT.Clone()
        $xmlNewTermLabel.Value = $label.Value.ToString()
        $xmlNewTermLabel.Language = $label.Language.ToString()
        $xmlNewTermLabel.IsDefaultForLanguage = $label.IsDefaultForLanguage.ToString()
        $xmlNewTerm.Labels.AppendChild($xmlNewTermLabel) | Out-Null
     }

     #Append new Term to Parent
     $parent.Terms.AppendChild($xmlNewTerm) | Out-Null

     Write-Host "Adding Term " -NoNewline
     Write-Host $_.name -ForegroundColor Green -NoNewline
     Write-Host " to Parent " -NoNewline
     Write-Host $parent.Name -ForegroundColor Green

     #If this term has child terms we need to loop through those
     if($_.Terms.Count -gt 0){
        #Recursively call itself
        Get-Terms $spContext $xmlNewTerm $_.Terms $xml -ExcludeIDs:$ExcludeIDs     
     }
 }
}

function Get-Groups($spContext, $groups, $xml, $groupToExport, [switch] $ExcludeIDs){

	 #Loop through all groups, ignoring system Groups
	 $groups | Where-Object { $_.IsSystemGroup -eq $false} | ForEach-Object{
   
	   #Check if we are getting groups or just group.
	   if($groupToExport -ne "")
	   {
		 if($groupToExport -ne $_.name){
		  #Return acts like a continue in ForEach-Object
		  return;
		 }
	   }
    
		#Add each group to export xml by cloning the template group,
		#populating it and appending it
		$xmlNewGroup = $global:xmlGroupT.Clone()
		$xmlNewGroup.Name = $_.name
		if (-Not $ExcludeIDs.IsPresent)
		{
			$xmlNewGroup.id = $_.id.ToString()
		}
		$xmlNewGroup.Description = $_.description
		if ($_.IsSiteCollectionGroup)
		{
			$xmlNewGroup.IsSiteCollectionGroup = "True"
		}
		$xml.TermStores.TermStore.Groups.AppendChild($xmlNewGroup) | Out-Null

		write-Host "Adding Group " -NoNewline
		write-Host $_.name -ForegroundColor Green

		$spContext.Load($_.TermSets)
		try
		{
			$spContext.ExecuteQuery()
		}
		catch
		{
		  Write-host "Error while loaded TermSets for Group " $xmlNewGroup.Name " " $_.Exception.Message -ForegroundColor Red
		  exit 1
		}

		Get-TermSets $spContext $xmlNewGroup $_.Termsets $xml -ExcludeIDs:$ExcludeIDs
	 }
}

function ExportTaxonomy($spContext, $termStore, $xml, $groupToExport, $path, $saveFileName, [switch] $ExcludeIDs){
   
	   $spContext.Load($termStore.Groups)
	   try
	   {
		 $spContext.ExecuteQuery();
	   }
	   catch
	   {
		 Write-host "Error while loaded Groups from TermStore " $_.Exception.Message -ForegroundColor Red
		 exit 1
	   }

   
	   Get-Groups $spContext $termStore.Groups $xml $groupToExport -ExcludeIDs:$ExcludeIDs

	   #Clean up empty tags/nodes
	   Clean-Template $xml

	   #Save file.
	   try
	   {
		   $xml.Save($path + "\NewTaxonomy.xml")
   

		   #Clean up empty <Term> unable to work out in Clean-Template.
		   Get-Content ($path + "\NewTaxonomy.xml") | Foreach-Object { $_ -replace "<Term><\/Term>", "" } | Set-Content ($path + "\" + $saveFileName)
		   Write-Host "Saving XML file " $saveFileName " to " $path

		   #Remove temp file
		   Remove-Item($path + "\Template.xml");
		   Remove-Item($path + "\NewTaxonomy.xml");
	   }
	   catch
	   {
			Write-host "Error saving XML File to disk " $_.Exception.Message -ForegroundColor Red
			exit 1
	   }
}
