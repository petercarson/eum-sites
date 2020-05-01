import { HttpClient } from "@microsoft/sp-http";
export interface ISitesListWebPartProps {
  webpartTitle: string;
  displayMode: string;
  groupBy: string;
  parentSiteURL: string;
  masterSiteURL: string;
  siteListName: string;
  currentWebRelativeUrl: string;
  currentWebAbsoluteUrl: string;
  siteMetadataSearchQuery: string;
  siteMetadataManagedProperties: string;
  displayUserSites: boolean;
  displayAvailableSites: boolean;
  displayAllSites: boolean;
  siteProvisioningApiUrl: string;
  tabHeaderUserSites: string;
  tabHeaderAvailableSites: string;
  tabHeaderAllSites: string;
  HttpClient: HttpClient;
  accessToken: string;
}
