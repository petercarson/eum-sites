import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, Environment, EnvironmentType } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  PropertyPaneToggle
} from '@microsoft/sp-webpart-base';

import "@pnp/polyfill-ie11";
import { sp } from "@pnp/sp";
import { StorageEntity, Web } from "@pnp/sp";
import * as strings from 'SitesListWebPartWebPartStrings';
import SitesListWebPart from './components/SitesListWebPart';
import { ISitesListWebPartProps } from './components/ISitesListWebPartProps';
import { AadTokenProvider, HttpClient, HttpClientResponse } from '@microsoft/sp-http';

export interface ISitesListWebPartWebPartProps {
  webpartTitle: string;
  parentSiteURL: string;
  displayMode: string;
  groupBy: string;
  masterSiteURL: string;
  siteListName: string;
  siteMetadataSearchQuery: string;
  siteMetadataManagedProperties: string;
  displayUserSites: boolean;
  displayAvailableSites: boolean;
  displayAllSites: boolean;
  siteProvisioningApiUrl: string;
  siteProvisioningApiClientID: string;
  tabHeaderUserSites: string;
  tabHeaderAvailableSites: string;
  tabHeaderAllSites: string;
  AadTokenProvider: AadTokenProvider;
  accessToken: string;
}

export default class SitesListWebPartWebPart extends BaseClientSideWebPart<ISitesListWebPartWebPartProps> {

  public async render(): Promise<void> {

    if (!(Environment.type === EnvironmentType.Local)) {
      await this.GetTenantProperties();

      if (this.properties.siteProvisioningApiUrl && this.properties.siteProvisioningApiClientID) {
        this.properties.AadTokenProvider = await this.context.aadTokenProviderFactory.getTokenProvider();
        await this.properties.AadTokenProvider.getToken(this.properties.siteProvisioningApiClientID).then((accessToken: string): void => {
          this.properties.accessToken = accessToken;
        });
      }
    }

    if (!this.properties.displayMode) {
      this.properties.displayMode = strings.DisplayModeTabs;
    }

    if (this.properties.displayUserSites == undefined) {
      this.properties.displayUserSites = true;
    }

    if (this.properties.displayAvailableSites == undefined) {
      this.properties.displayAvailableSites = true;
    }

    if (this.properties.displayAllSites == undefined) {
      this.properties.displayAllSites = true;
    }

    const element: React.ReactElement<ISitesListWebPartProps> = React.createElement(
      SitesListWebPart,
      {
        webpartTitle: this.properties.webpartTitle,
        displayMode: this.properties.displayMode,
        groupBy: this.properties.groupBy,
        parentSiteURL: this.properties.parentSiteURL,
        masterSiteURL: this.properties.masterSiteURL,
        siteListName: this.properties.siteListName,
        currentWebRelativeUrl: this.context.pageContext.web.serverRelativeUrl,
        currentWebAbsoluteUrl: this.context.pageContext.web.absoluteUrl,
        siteMetadataSearchQuery: this.properties.siteMetadataSearchQuery,
        siteMetadataManagedProperties: this.properties.siteMetadataManagedProperties,
        displayUserSites: this.properties.displayUserSites,
        displayAvailableSites: this.properties.displayAvailableSites,
        displayAllSites: this.properties.displayAllSites,
        siteProvisioningApiUrl: this.properties.siteProvisioningApiUrl,
        tabHeaderUserSites: this.properties.tabHeaderUserSites,
        tabHeaderAvailableSites: this.properties.tabHeaderAvailableSites,
        tabHeaderAllSites: this.properties.tabHeaderAllSites,
        HttpClient: this.context.httpClient,
        accessToken: this.properties.accessToken,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  public onInit(): Promise<void> {

    return super.onInit().then(_ => {
      sp.setup({
        sp: {
          headers: {
            Accept: "application/json;odata=verbose",
          },
          baseUrl: this.GetMasterSiteAbsoluteUrl() // the lists are in the master site so all requests should go there
        },
      });
    });
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
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
                  label: 'Web Part Title'
                }),
                PropertyPaneTextField('parentSiteURL', {
                  label: strings.ParentSiteURLFieldLabel
                }),
                PropertyPaneDropdown('displayMode', {
                  label: strings.DisplayModeFieldLabel,
                  options: [
                    { key: strings.DisplayModeTabs, text: strings.DisplayModeTabs },
                    { key: strings.DisplayModeListing, text: strings.DisplayModeListing },
                    { key: strings.DisplayModeTiles, text: strings.DisplayModeTiles },
                    { key: strings.DisplayModeAuto, text: strings.DisplayModeAuto }
                  ],
                  selectedKey: strings.DisplayModeTabs
                }),
                PropertyPaneDropdown('groupBy', {
                  label: strings.GroupByFieldLabel,
                  options: [
                    { key: strings.GroupByTitle, text: strings.GroupByTitle },
                    { key: strings.GroupByParent, text: strings.GroupByParent }
                  ],
                  selectedKey: strings.GroupByTitle
                }),
                PropertyPaneTextField('masterSiteURL', {
                  label: strings.MasterSiteURLFieldLabel
                }),
                PropertyPaneTextField('siteListName', {
                  label: strings.SiteListNameFieldLabel
                }),
                PropertyPaneTextField('tabHeaderUserSites', {
                  label: strings.TabHeaderUserSitesFieldLabel
                }),
                PropertyPaneTextField('tabHeaderAvailableSites', {
                  label: strings.TabHeaderAvailableSitesFieldLabel
                }),
                PropertyPaneTextField('tabHeaderAllSites', {
                  label: strings.TabHeaderAllSitesFieldLabel
                })
              ]
            },
          ]
        },
        {
          header: {
            description: strings.SearchSettingsPaneDescription
          },
          groups: [
            {
              groupName: strings.SearchSettingsGroupName,
              groupFields: [
                PropertyPaneToggle('displayUserSites', {
                  label: strings.DisplayUserSites,
                  onText: strings.ToggleOnText,
                  offText: strings.ToggleOffText,
                  checked: true
                }),
                PropertyPaneToggle('displayAvailableSites', {
                  label: strings.DisplayAvailableSites,
                  onText: strings.ToggleOnText,
                  offText: strings.ToggleOffText,
                  checked: true
                }),
                PropertyPaneToggle('displayAllSites', {
                  label: strings.DisplayAllSites,
                  onText: strings.ToggleOnText,
                  offText: strings.ToggleOffText,
                  checked: true
                }),
                PropertyPaneTextField('siteMetadataSearchQuery', {
                  label: strings.SiteMetadataSearchQueryFieldLabel,
                  multiline: true
                }),
                PropertyPaneTextField('siteMetadataManagedProperties', {
                  label: strings.SiteMetadataManagedPropertiesFieldLabel,
                  description: strings.SiteMetadataManagedPropertiesFieldDescription,
                  multiline: true
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
            },
          ]
        }
      ]
    };
  }

  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }

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
    }
  }

  // **************************************
  // Private Methods - Url Parsing Helpers
  // **************************************
  private GetCurrentWebAbsoluteUrl(): string {
    return this.context.pageContext.web.absoluteUrl;
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
