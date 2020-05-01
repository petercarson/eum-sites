define([], function() {
  return {
    "PropertyPaneDescription": "Enter the Parent URL (relative or full) in the Parent URL field. The web part will render a listing of all its child sites. If no Parent URL is provided, the current site URL will be used.",
    "BasicGroupName": "Site Listing Settings",
    "SearchSettingsPaneDescription": "Search query settings for retrieving site metadata from the lists in each site collection. This is used to display sites that the user has access to.",
    "SiteMetadataSearchSettingsGroupName": "Site Metadata Search Settings",

    "ToggleOnText": "Yes",
    "ToggleOffText": "No",
    
    "DisplayUserSites": "Display sites user has access to",
    "DisplayAvailableSites": "Display available sites for the user",
    "DisplayAllSites": "Display all sites",

    "MasterSiteURLFieldLabel": "Master Site URL",
    "SiteListNameFieldLabel": "Site List Name",
    "ParentSiteURLFieldLabel": "Parent URL",
    "DisplayModeFieldLabel": "Display Mode",
    "DisplayModeAuto": "Auto",
    "DisplayModeListing": "List",
    "DisplayModeTabs": "Tabs",
    "DisplayModeTiles": "Tiles",
    "GroupByFieldLabel": "Group By",
    "GroupByTitle": "Title",
    "GroupByParent": "Division",

    "SiteMetadataSearchQueryFieldLabel": "Site Metadata Search Query",
    "SiteMetadataManagedPropertiesFieldLabel": "Site Metadata Managed Properties",
    "SiteMetadataManagedPropertiesFieldDescription": "Enter a comma-separated list of managed properties to retrieve.",

    "SiteProvisioningApiUrlFieldLabel" : "Site provisioning API URL",
    "SiteProvisioningApiUrlFieldDescription": "Enter the site provisioning API URL used to retrieve sites the user does not have access to, and all sites. If no URL is provided, webpart will directly query the Sites list via SharePoint API to retrieve all sites.",

    "SiteProvisioningApiClientIDFieldLabel": "Site Provisioning API Azure AD Client ID",
    "TenantPropertyDescription": "The web part will retrieve this value from your tenant properties. Enter a value in this field if you wish to override the tenant property.",

    "TabHeaderUserSitesFieldLabel": "Tab title for current user's sites",
    "TabHeaderAvailableSitesFieldLabel" : "Tab title for sites available for current user",
    "TabHeaderAllSitesFieldLabel" : "Tab title for all sites",

    "NoSitesFoundText": "No sites available.",
    "LoadingText": "Loading..."
  }
});