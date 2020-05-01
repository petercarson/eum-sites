declare interface ISitesListWebPartWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  SearchSettingsPaneDescription: string;
  SearchSettingsGroupName: string;

  ToggleOnText: string;
  ToggleOffText: string;

  DisplayUserSites: string;
  DisplayAvailableSites: string;
  DisplayAllSites: string;

  MasterSiteURLFieldLabel: string;
  SiteListNameFieldLabel: string;
  ParentSiteURLFieldLabel: string;
  DisplayModeFieldLabel: string;
  DisplayModeAuto: string;
  DisplayModeListing: string;
  DisplayModeTabs: string;
  DisplayModeTiles: string;
  GroupByFieldLabel: string;
  GroupByTitle: string;
  GroupByParent: string;

  SiteMetadataSearchQueryFieldLabel: string; 
  SiteMetadataManagedPropertiesFieldLabel: string; 
  SiteMetadataManagedPropertiesFieldDescription: string;

  SiteProvisioningApiUrlFieldLabel: string;
  SiteProvisioningApiUrlFieldDescription: string;

  TabHeaderUserSitesFieldLabel: string;
  TabHeaderAvailableSitesFieldLabel: string;
  TabHeaderAllSitesFieldLabel: string;

  NoSitesFoundText: string;
  LoadingText: string;
}

declare module 'SitesListWebPartWebPartStrings' {
  const strings: ISitesListWebPartWebPartStrings;
  export = strings;
}
