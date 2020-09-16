import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, Environment, EnvironmentType } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'SiteRequestFormWebPartStrings';
import SiteRequestForm from './components/SiteRequestForm';
import { ISiteRequestFormProps } from './components/ISiteRequestFormProps';
import { AadTokenProvider } from '@microsoft/sp-http';
import * as microsoftTeams from '@microsoft/teams-js';

// IE Support
import "@pnp/polyfill-ie11";

import { sp, StorageEntity, Web } from "@pnp/sp";

export interface ISiteRequestFormWebPartProps {
  webpartTitle: string;
  description: string;
  masterSiteURL: string;
  siteListName: string;
  divisionsListName: string;
  siteTemplatesListName: string;
  blacklistedWordsListName: string;
  defaultNewItemUrl: string;
  titleFieldLabel: string;
  divisionFieldLabel: string;
  siteTemplateFieldLabel: string;
  preselectedDivision: string;
  preselectedSiteTemplate: string;
  siteProvisioningApiUrl: string;
  siteProvisioningApiClientID: string;
  AadTokenProvider: AadTokenProvider;
  siteProvisioningExternalSharingPrefix : string;
}

export default class SiteRequestFormWebPart extends BaseClientSideWebPart<ISiteRequestFormWebPartProps> {
  private _teamsContext: microsoftTeams.Context;

  public async render(): Promise<void> {
    if (!(Environment.type === EnvironmentType.Local)) {
      await this.GetTenantProperties();
      if (this.properties.siteProvisioningApiUrl && this.properties.siteProvisioningApiClientID) {
        this.properties.AadTokenProvider = await this.context.aadTokenProviderFactory.getTokenProvider();
      }
    }


    const element: React.ReactElement<ISiteRequestFormProps> = React.createElement(
      SiteRequestForm,
      {
        webpartTitle: this.properties.webpartTitle,
        description: this.properties.description,
        divisionsListName: this.properties.divisionsListName,
        siteTemplatesListName: this.properties.siteTemplatesListName,
        sitesListName: this.properties.siteListName,
        blacklistedWordsListName: this.properties.blacklistedWordsListName,
        titleFieldLabel: this.properties.titleFieldLabel,
        divisionFieldLabel: this.properties.divisionFieldLabel,
        siteTemplateFieldLabel: this.properties.siteTemplateFieldLabel,
        preselectedDivision: this.properties.preselectedDivision,
        preselectedSiteTemplate: this.properties.preselectedSiteTemplate,
        siteProvisioningApiUrl: this.properties.siteProvisioningApiUrl,
        context: this.context,
        tenantUrl: this.GetRootSiteUrl(),
        HttpClient: this.context.httpClient,
        siteProvisioningApiClientID: this.properties.siteProvisioningApiClientID,
        AadTokenProvider: this.properties.AadTokenProvider,
        siteProvisioningExternalSharingPrefix : this.properties.siteProvisioningExternalSharingPrefix
      }
    );
    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected onInit(): Promise<any> {
    return super.onInit().then(_ => {
      // initialize PnP-JS
      sp.setup({
        sp: {
          headers: {
            Accept: "application/json;odata=verbose",
          },
          baseUrl: this.GetMasterSiteAbsoluteUrl()
        },
      });

      // determine if teams context or SP context
      let retVal: Promise<any> = Promise.resolve();
      if (this.context.microsoftTeams) {
        retVal = new Promise((resolve, reject) => {
          this.context.microsoftTeams.getContext(context => {
            this._teamsContext = context;
            resolve();
          });
        });
      }
      return retVal;
    });
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('webpartTitle', {
                  label: strings.WebPartTitleFieldLabel
                }),
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel,
                  multiline: true
                }),
                PropertyPaneTextField('masterSiteURL', {
                  label: strings.MasterSiteURLFieldLabel
                }),
                PropertyPaneTextField('siteListName', {
                  label: strings.SiteListNameFieldLabel
                }),
                PropertyPaneTextField('divisionsListName', {
                  label: strings.DivisionsListNameFieldLabel
                }),
                PropertyPaneTextField('siteTemplatesListName', {
                  label: strings.SiteTemplatesListNameFieldLabel
                }),
                PropertyPaneTextField('blacklistedWordsListName', {
                  label: strings.BlacklistedWordsListNameFieldLabel
                }),
                PropertyPaneTextField('defaultNewItemUrl', {
                  label: strings.DefaultNewItemUrl
                }),
                PropertyPaneTextField('divisionFieldLabel', {
                  label: strings.DivisionFieldLabel,
                  description: strings.DivisionFieldLabelDescription
                }),   
                PropertyPaneTextField('siteTemplateFieldLabel', {
                  label: strings.SiteTemplateFieldLabel,
                  description: strings.SiteTemplateFieldLabelDescription
                }),       
                PropertyPaneTextField('titleFieldLabel', {
                  label: strings.TitleFieldLabel,
                  description: strings.TitleFieldLabelDescription
                }),        
                PropertyPaneTextField('preselectedDivision', {
                  label: strings.PreselectedDivisionLabel,
                  description: strings.PreselectedDivisionDescription
                }),
                PropertyPaneTextField('preselectedSiteTemplate', {
                  label: strings.PreselectedSiteTemplateLabel,
                  description: strings.PreselectedSiteTemplateDescription
                }),
                PropertyPaneTextField('siteProvisioningApiUrl', {
                  label: strings.SiteProvisioningApiUrlFieldLabel,
                  description: strings.SiteProvisioningApiUrlFieldDescription,
                }),
                PropertyPaneTextField('siteProvisioningApiClientID', {
                  label: strings.SiteProvisioningApiClientIDFieldLabel,
                  description: strings.TenantPropertyDescription
                })
              ]
            }
          ]
        }
      ]
    };
  }


  // **************************************
  // Private Methods
  // **************************************
  private async GetTenantProperties(): Promise<void> {
    if (Environment.type != EnvironmentType.Local) {
      // get the tenant properties if they are not set in the web part
      let currentWeb: Web = new Web(this.GetCurrentWebAbsoluteUrl());

      let storageEntity: StorageEntity = null;
      if (!this.properties.siteProvisioningApiUrl) {
        storageEntity = await currentWeb.getStorageEntity('siteProvisioningApiUrl');
        this.properties.siteProvisioningApiUrl = storageEntity ? storageEntity.Value : null;
      }

      if (!this.properties.siteProvisioningApiClientID) {
        storageEntity = await currentWeb.getStorageEntity('siteProvisioningApiClientID');
        this.properties.siteProvisioningApiClientID = storageEntity ? storageEntity.Value : null;
      }
      if (!this.properties.siteProvisioningExternalSharingPrefix) {
        storageEntity = await currentWeb.getStorageEntity('siteProvisioningExternalSharingPrefix');
        this.properties.siteProvisioningExternalSharingPrefix = storageEntity ? storageEntity.Value : null;
      }
    }
  }

  private GetCurrentWebAbsoluteUrl(): string {
    if (this._teamsContext) {
      return this._teamsContext.teamSiteUrl;
    } else {
      return this.context.pageContext.web.absoluteUrl;
    }
  }

  private GetRootSiteUrl(): string {
    return (new URL(this.GetCurrentWebAbsoluteUrl())).origin;
  }

  private GetMasterSiteAbsoluteUrl(): string {
    let masterSiteAbsoluteUrl: string = this.properties.masterSiteURL;
    if (masterSiteAbsoluteUrl.charAt(0) === '/') {
      masterSiteAbsoluteUrl = this.GetRootSiteUrl() + masterSiteAbsoluteUrl;
    }
    return masterSiteAbsoluteUrl;
  }
}
