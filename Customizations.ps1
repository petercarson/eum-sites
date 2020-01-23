function ApplySiteCustomizations () {
	Param
	(
		[Parameter (Mandatory = $true)][int]$listItemID
	)

	$pendingSiteCollection = Get-PnPListItem -List $SiteListName -Query "
	<View>
		<Query>
			<Where>
				<Eq>
					<FieldRef Name='ID'/>
					<Value Type='Integer'>$listItemID</Value>
				</Eq>
			</Where>
		</Query>
        <ViewFields>
            <FieldRef Name='ID'></FieldRef>
            <FieldRef Name='Title'></FieldRef>
            <FieldRef Name='EUMSiteURL'></FieldRef>
            <FieldRef Name='EUMAlias'></FieldRef>
            <FieldRef Name='EUMPublicGroup'></FieldRef>
            <FieldRef Name='EUMBreadcrumbHTML'></FieldRef>
            <FieldRef Name='EUMParentURL'></FieldRef>
            <FieldRef Name='EUMSiteTemplate'></FieldRef>
            <FieldRef Name='EUMDivision'></FieldRef>
            <FieldRef Name='EUMCreateTeam'></FieldRef>
            <FieldRef Name='EUMCreateOneNote'></FieldRef>
            <FieldRef Name='EUMCreatePlanner'></FieldRef>
            <FieldRef Name='Author'></FieldRef>
        </ViewFields>
	</View>"

	if ($pendingSiteCollection.Count -eq 1) {
		[string]$eumSiteTemplate = $pendingSiteCollection["EUMSiteTemplate"].LookupValue
		[string]$siteURL = $pendingSiteCollection["EUMSiteURL"]

		$siteTemplate = Get-PnPListItem -List "Site Templates" -Query "
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
														</ViewFields>
													</View>"
		
		$baseSiteTemplate = ""
		$baseSiteType = ""
		$pnpSiteTemplate = ""

		if ($siteTemplate.Count -eq 1) {
			$baseSiteTemplate = $siteTemplate["BaseClassicSiteTemplate"]
			$baseSiteType = $siteTemplate["BaseModernSiteType"]
			if ($siteTemplate["PnPSiteTemplate"] -ne $null) {
                $pnpSiteTemplate = "$pnpTemplatePath\$($siteTemplate["PnPSiteTemplate"].LookupValue)"
			}
		}
		
		if (($pnpSiteTemplate -like "*Client-Template.xml")) {
			Helper-Connect-PnPOnline -Url $siteURL

			Add-PnPFolder -Name "Quotes" -Folder "/Shared Documents"
			Add-PnPFolder -Name "Signed Quotes" -Folder "/Shared Documents/Quotes"
			Add-PnPFolder -Name "Invoices" -Folder "/Shared Documents"

			Add-PnPFolder -Name "Business Development" -Folder "/Private Documents"
			Add-PnPFolder -Name "Confidential" -Folder "/Private Documents"
			Add-PnPFolder -Name "Quotes" -Folder "/Private Documents"

			Remove-PnPContentTypeFromList -List "Shared Documents" -ContentType "Document"
			Remove-PnPContentTypeFromList -List "Private Documents" -ContentType "Document"
		}
	}
}

function ApplyChannelCustomizations () {
	Param
	(
		[Parameter (Mandatory = $true)][int]$listItemID
	)
}
