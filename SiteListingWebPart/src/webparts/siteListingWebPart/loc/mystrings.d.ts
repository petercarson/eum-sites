declare interface ISiteListingWebPartWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;

  MasterSiteURLFieldLabel: string;
  SiteListNameFieldLabel: string;
  ParentSiteURLFieldLabel: string;
  DisplayModeFieldLabel: string;
  DisplayModeAuto: string;
  DisplayModeListing: string;
  DisplayModeTabs: string;
  IncludeBootstrapFieldLabel: string;
  NoSitesFoundText: string;
}

declare module 'SiteListingWebPartWebPartStrings' {
  const strings: ISiteListingWebPartWebPartStrings;
  export = strings;
}
