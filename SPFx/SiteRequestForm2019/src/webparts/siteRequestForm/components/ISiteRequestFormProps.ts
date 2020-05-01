import { WebPartContext } from "@microsoft/sp-webpart-base";
import { HttpClient } from '@microsoft/sp-http';

export interface ISiteRequestFormProps {
  webpartTitle?: string;
  description?: string;
  divisionsListName: string;
  siteTemplatesListName: string;
  sitesListName: string;
  context: WebPartContext;
  tenantUrl: string;
  titleFieldLabel: string;
  preselectedDivision: string;
  preselectedSiteTemplate: string;
  siteProvisioningApiUrl: string;
  HttpClient: HttpClient;
}
