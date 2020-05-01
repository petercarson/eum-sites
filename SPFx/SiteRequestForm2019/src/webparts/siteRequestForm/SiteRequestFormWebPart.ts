import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'SiteRequestFormWebPartStrings';
import SiteRequestForm from './components/SiteRequestForm';
import { ISiteRequestFormProps } from './components/ISiteRequestFormProps';

import { sp } from "@pnp/sp";
import "core-js/es6/array";
import "es6-map/implement";
import "core-js/modules/es6.array.find";

export interface ISiteRequestFormWebPartProps {
  webpartTitle: string;
  description: string;
  masterSiteURL: string;
  siteListName: string;
  divisionsListName: string;
  siteTemplatesListName: string;
  defaultNewItemUrl: string;
  titleFieldLabel: string;
  preselectedDivision: string;
  preselectedSiteTemplate: string;
  siteProvisioningApiUrl: string;
}

export default class SiteRequestFormWebPart extends BaseClientSideWebPart<ISiteRequestFormWebPartProps> {

  public render(): void {

    const element: React.ReactElement<ISiteRequestFormProps> = React.createElement(
      SiteRequestForm,
      {
        webpartTitle: this.properties.webpartTitle,
        description: this.properties.description,
        divisionsListName: this.properties.divisionsListName,
        siteTemplatesListName: this.properties.siteTemplatesListName,
        sitesListName: this.properties.siteListName,
        titleFieldLabel: this.properties.titleFieldLabel,
        preselectedDivision: this.properties.preselectedDivision,
        preselectedSiteTemplate: this.properties.preselectedSiteTemplate,
        siteProvisioningApiUrl: this.properties.siteProvisioningApiUrl,
        context: this.context,
        tenantUrl: this.GetRootSiteUrl(),
        HttpClient: this.context.httpClient
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
                PropertyPaneTextField('defaultNewItemUrl', {
                  label: strings.DefaultNewItemUrl
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
