import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'TeamsChannelWebPartStrings';
import TeamsChannel from './components/TeamsChannel';
import { ITeamsChannelProps } from './components/ITeamsChannelProps';
import * as microsoftTeams from '@microsoft/teams-js';
import { sp } from "@pnp/sp";

export interface ITeamsChannelWebPartProps {
  WebPartTitle: string;
  Description: string;
  MasterSiteUrl: string;
  SitesListName: string;
  DivisionsListName: string;
  TeamsChannelsListName: string;
  ChannelTemplatesListName: string;
}

export default class TeamsChannelWebPart extends BaseClientSideWebPart<ITeamsChannelWebPartProps> {

  private _teamsContext: microsoftTeams.Context;

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

  public render(): void {
    const element: React.ReactElement<ITeamsChannelProps> = React.createElement(
      TeamsChannel,
      {
        WebPartTitle: this.properties.WebPartTitle,
        Description: this.properties.Description,
        MasterSiteUrl: this.GetMasterSiteAbsoluteUrl(),
        SitesListName: this.properties.SitesListName,
        TeamsChannelsListName: this.properties.TeamsChannelsListName,
        ChannelTemplatesListName: this.properties.ChannelTemplatesListName,
        DivisionsListName: this.properties.DivisionsListName,
        SiteName: this.GetCurrentWebTitle(),
        SiteUrl: this.GetCurrentWebAbsoluteUrl()
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
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
                PropertyPaneTextField('WebPartTitle', {
                  label: strings.WebPartTitleFieldLabel
                }),
                PropertyPaneTextField('Description', {
                  label: strings.DescriptionFieldLabel,
                  multiline: true
                }),
                PropertyPaneTextField('MasterSiteUrl', {
                  label: strings.MasterSiteUrlFieldLabel
                }),
                PropertyPaneTextField('SitesListName', {
                  label: strings.SitesListNameFieldLabel
                }),
                PropertyPaneTextField('DivisionsListName', {
                  label: strings.DivisionsListNameFieldLabel
                }),
                PropertyPaneTextField('TeamsChannelsListName', {
                  label: strings.TeamsChannelsListNameFieldLabel
                }),
                PropertyPaneTextField('ChannelTemplatesListName', {
                  label: strings.ChannelTemplatesListNameFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }

  private GetCurrentWebAbsoluteUrl(): string {
    if (this._teamsContext) {
      return this._teamsContext.teamSiteUrl;
    } else {
      return this.context.pageContext.web.absoluteUrl;
    }
  }

  private GetCurrentWebTitle(): string {
    if (this._teamsContext) {
      return this._teamsContext.teamName;
    } else {
      return this.context.pageContext.web.title;
    }
  }

  private GetRootSiteUrl(): string {
    return (new URL(this.GetCurrentWebAbsoluteUrl())).origin;
  }

  private GetMasterSiteAbsoluteUrl(): string {
    let masterSiteAbsoluteUrl: string = this.properties.MasterSiteUrl;
    if (masterSiteAbsoluteUrl.charAt(0) === '/') {
      masterSiteAbsoluteUrl = this.GetRootSiteUrl() + masterSiteAbsoluteUrl;
    }
    return masterSiteAbsoluteUrl;
  }
}
