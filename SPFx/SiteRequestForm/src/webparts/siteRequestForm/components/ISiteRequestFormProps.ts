import { WebPartContext } from "@microsoft/sp-webpart-base";
import { HttpClient } from '@microsoft/sp-http';
import { AadTokenProvider } from '@microsoft/sp-http';

export interface ISiteRequestFormProps {
  webpartTitle?: string;
  description?: string;
  divisionsListName: string;
  siteTemplatesListName: string;
  sitesListName: string;
  blacklistedWordsListName: string;
  context: WebPartContext;
  tenantUrl: string;
  titleFieldLabel: string;
  divisionFieldLabel: string;
  siteTemplateFieldLabel: string;
  preselectedDivision: string;
  preselectedSiteTemplate: string;
  siteProvisioningApiUrl: string;
  siteProvisioningApiClientID: string;
  HttpClient: HttpClient;
  AadTokenProvider: AadTokenProvider;
  siteProvisioningExternalSharingPrefix : string;
}