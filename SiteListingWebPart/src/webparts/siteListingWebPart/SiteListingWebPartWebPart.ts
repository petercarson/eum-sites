import { Version, Environment, EnvironmentType } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  PropertyPaneCheckbox
} from '@microsoft/sp-webpart-base';
import { SPComponentLoader } from '@microsoft/sp-loader';

import { sp } from "@pnp/pnpjs";
import { ISiteListItem } from './ISiteListItem';
import styles from './SiteListingWebPartWebPart.module.scss';
import * as strings from 'SiteListingWebPartWebPartStrings';

export interface ISiteListingWebPartWebPartProps {
  parentSiteURL: string;
  displayMode: string;
  masterSiteURL: string;
  siteListName: string;
  includeBootstrap: boolean;
}

export default class SiteListingWebPartWebPart extends BaseClientSideWebPart<ISiteListingWebPartWebPartProps> {

  private divContainerClass: string = "SiteListing";
  private siteRequestTemplate: Document = null;

  public render(): void {
    this.LoadJQueryAndBootstrap();

    // base container for the A-Z listing
    this.domElement.innerHTML = `<div class="${this.divContainerClass} ${styles.siteListingWebPart}"></div>`;

    this.GetListItemsAndRenderListing();
  }

  protected onInit(): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error?: any) => void): void => {
      sp.setup({
        sp: {
          headers: {
            "Accept": "application/json; odata=verbose"
          },
          baseUrl: this.GetMasterSiteAbsoluteUrl() // the list is in the master site so all requests should go there
        },
      });
      resolve();
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
                PropertyPaneTextField('parentSiteURL', {
                  label: strings.ParentSiteURLFieldLabel
                }),
                PropertyPaneDropdown('displayMode', {
                  label: strings.DisplayModeFieldLabel,
                  options: [
                    { key: strings.DisplayModeAuto, text: strings.DisplayModeAuto },
                    { key: strings.DisplayModeListing, text: strings.DisplayModeListing },
                    { key: strings.DisplayModeTabs, text: strings.DisplayModeTabs }
                  ],
                  selectedKey: strings.DisplayModeAuto
                }),
                PropertyPaneTextField('masterSiteURL', {
                  label: strings.MasterSiteURLFieldLabel
                }),
                PropertyPaneTextField('siteListName', {
                  label: strings.SiteListNameFieldLabel
                }),
                PropertyPaneCheckbox('includeBootstrap', {
                  text: strings.IncludeBootstrapFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }

  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }


  // ***********************************
  // Private Methods - List Operations
  // ***********************************
  private GetListItemsAndRenderListing(): void {
    if (Environment.type === EnvironmentType.Local) {
      // load dummy data
      this.GetSiteListItemsDebug();
    } else {
      this.GetSiteListItems();
    }
  }

  private GetSiteListItems(): void {
    let EUMParentURL: string = "";
    if (this.properties.parentSiteURL) {
      // parentSiteURL is set so filter the site list by its children
      EUMParentURL = this.GetParentSiteRelativeUrl();
    } else {
      // no parentSiteURL provided, so assume current site
      EUMParentURL = this.GetCurrentWebRelativeUrl();
    }

    if (EUMParentURL.toLocaleLowerCase() == this.properties.masterSiteURL.toLocaleLowerCase()) {
      this.GetAllSiteListItems();
    } else {
      this.GetSiteListItemsFiltered(EUMParentURL);
    }
  }

  private GetSiteListItemsFiltered(EUMParentURL: string): void {
    // get all children of the parent url
    let viewXml: string = `<View>
        <Query>
          <Where>
            <Or>
              <Eq><FieldRef Name="EUMParentURL"/><Value Type="URL">${EUMParentURL}</Value></Eq>
              <Eq><FieldRef Name="EUMParentURL"/><Value Type="URL">${EUMParentURL}</Value></Eq>
            </Or>
          </Where>
          <OrderBy>
            <FieldRef Name="Title" Ascending="TRUE"/>
          </OrderBy>
        </Query>
        <ViewFields>
          <FieldRef Name="ID"></FieldRef>
          <FieldRef Name="Title"></FieldRef>
          <FieldRef Name="EUMSiteURL"></FieldRef>
          <FieldRef Name="EUMParentURL"></FieldRef>
        </ViewFields>
      </View>`;

    sp.web
      .lists.getByTitle(this.properties.siteListName)
      .usingCaching()
      .getItemsByCAMLQuery({ ViewXml: viewXml })
      .then((items: ISiteListItem[]): void => {
        this.RenderSiteListing(items);
      }, (error: any): void => {
        this.PrintErrorMessage(`Failed loading sites. Error: ${error}`);
      });
  }

  private GetAllSiteListItems(): void {
    sp.web
      .lists.getByTitle(this.properties.siteListName)
      .items.select('Id, Title, EUMSiteURL, EUMParentURL')
      .top(5000)
      .orderBy('Title', true)
      .usingCaching()
      .get()
      .then((items: ISiteListItem[]): void => {
        this.RenderSiteListing(items);
      }, (error: any): void => {
        this.PrintErrorMessage(`Failed loading sites. Error: ${error}`);
      });
  }

  // **************************************
  // Private Methods - Url Parsing Helpers
  // **************************************
  private GetCurrentWebAbsoluteUrl(): string {
    return this.context.pageContext.web.absoluteUrl;
  }

  private GetCurrentWebRelativeUrl(): string {
    return this.context.pageContext.web.serverRelativeUrl;
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

  private GetParentSiteRelativeUrl(): string {
    return this.properties.parentSiteURL.replace(this.GetRootSiteUrl(), "");
  }

  // ***********************************
  // Private Methods - DOM Helpers
  // ***********************************
  private PrintErrorMessage(errMsg: string): void {
    let html: string = `<div class="alert alert-danger" id="az-error"><p>${errMsg}</div>`;
    this.domElement.innerHTML = html;
  }

  private RenderSiteListing(items: ISiteListItem[]): void {

    if (items.length == 0) {
      this.RenderNoSitesMessage();
    } else if (this.properties.displayMode === strings.DisplayModeTabs) {
      this.RenderSitesTabs(items);
    } else if (this.properties.displayMode === strings.DisplayModeListing) {
      this.RenderSitesList(items);
    } else {
      this.RenderAuto(items);
    }
  }

  private RenderNoSitesMessage(): void {
    this.domElement.querySelector(`.${this.divContainerClass}`).innerHTML = `
    <div>
      <p>
        ${strings.NoSitesFoundText}
      </p>
    </div>`;
  }

  private RenderAuto(items: ISiteListItem[]): void {
    if (items.length == 1) {
      window.location.href = items[0].EUMSiteURL.Url;
    } else if (items.length >= 10) {
      this.RenderSitesTabs(items);
    } else if (items.length > 1 && items.length < 10) {
      this.RenderSitesList(items);
    }
  }

  private RenderSitesList(items: ISiteListItem[]): void {
    if (items.length > 0) {
      const listHTML: string[] = [];

      let groupedSites: Object = this.GroupSites(items);
      for (let key in groupedSites) {
        listHTML.push(groupedSites[key]);
      }

      this.domElement.querySelector(`.${this.divContainerClass}`).innerHTML = `
        <div class="siteslist">
          <ul>
            ${listHTML.join('')}
          </ul>
        </div>`;
    }
  }

  private RenderSitesTabs(items: ISiteListItem[]): void {
    if (items.length > 0) {
      const tabsHTML: string[] = [];
      const tabsContentHTML: string[] = [];

      let groupedSites: Object = this.GroupSites(items);
      for (let key in groupedSites) {
        let tabTitle: string = (key === "0-9" ? "&#35;" : key);
        let tabID: string = (key === "0-9" ? "num" : key);

        let activeTabClasses: string = tabsHTML.length === 0 ? `active` : "";
        tabsHTML.push(`<li class="${activeTabClasses}"><a data-toggle="tab" href="#${tabID}">${tabTitle}</a></li>`);

        let activeContentClasses: string = tabsContentHTML.length === 0 ? "active in" : "";
        tabsContentHTML.push(`<div id="${tabID}" class="tab-pane fade ${activeContentClasses}"><ul>${groupedSites[key]}</ul></div>`);
      }

      this.domElement.querySelector(`.${this.divContainerClass}`).innerHTML = `<ul class="nav nav-tabs ${styles.azTabs}">${tabsHTML.join('')}</ul>
      <div class="tab-content ${styles.azTabsContent}">${tabsContentHTML.join('')}</div>`;
    }
  }

  private GroupSites(items: ISiteListItem[]): Object {
    let sitesObject: Object = {};

    for (let i: number = 0, max = items.length; i < max; i++) {
      let currentSite: ISiteListItem = items[i];

      if (currentSite && currentSite.Title && currentSite.EUMSiteURL && currentSite.EUMSiteURL.Url) {
        let siteHTML: string = `<li><a href='${currentSite.EUMSiteURL.Url}'>${currentSite.Title}</a></li>`;

        let firstLetter: any = (currentSite.Title.charAt(0).toUpperCase());
        if (isNaN(firstLetter)) {
          if (sitesObject[firstLetter]) {
            sitesObject[firstLetter] += siteHTML;
          } else {
            sitesObject[firstLetter] = siteHTML;
          }
        } else {
          if (sitesObject["0-9"]) {
            sitesObject["0-9"] += siteHTML;
          } else {
            sitesObject["0-9"] = siteHTML;
          }
        }
      }
    }
    return sitesObject;
  }

  // ****************************
  // Private Methods - Debugging
  // ****************************
  private LoadJQueryAndBootstrap(): void {
    if (this.properties.includeBootstrap) {
      SPComponentLoader.loadScript('https://cdnjs.cloudflare.com/ajax/libs/jquery/3.3.1/jquery.min.js').then((jQuery: any): void => {
        SPComponentLoader.loadCss('https://cdnjs.cloudflare.com/ajax/libs/twitter-bootstrap/3.3.7/css/bootstrap.min.css');
        SPComponentLoader.loadScript('https://cdnjs.cloudflare.com/ajax/libs/twitter-bootstrap/3.3.7/js/bootstrap.min.js');
      });
    }
  }

  private GetSiteListItemsDebug(): void {
    this.GetMockItems()
      .then((items: ISiteListItem[]): void => {
        this.RenderSiteListing(items);
      });
  }

  private GetMockItems(): Promise<ISiteListItem[]> {
    return Promise.resolve([{
      Id: 0,
      Title: "A Site",
      EUMSiteURL: { "Description": "A site", "Url": "" },
      EUMParentURL: { "Description": "EIT Intranet", "Url": "https://envisionitdev.sharepoint.com/sites/eitintranet" }
    },
    {
      Id: 1,
      Title: "B Site",
      EUMSiteURL: { "Description": "B site", "Url": "https://envisionitdev.sharepoint.com/sites/bsite" },
      EUMParentURL: { "Description": "EIT Intranet", "Url": "https://envisionitdev.sharepoint.com/sites/eitintranet" }
    },
    {
      Id: 3,
      Title: "C Site",
      EUMSiteURL: { "Description": "C site", "Url": "https://envisionitdev.sharepoint.com/sites/csite" },
      EUMParentURL: { "Description": "EIT Intranet", "Url": "https://envisionitdev.sharepoint.com/sites/eitintranet" }
    },
    {
      Id: 4,
      Title: "1 Site",
      EUMSiteURL: { "Description": "1 site", "Url": "https://envisionitdev.sharepoint.com/sites/numsite" },
      EUMParentURL: { "Description": "EIT Intranet", "Url": "https://envisionitdev.sharepoint.com/sites/eitintranet" }
    }, {
      Id: 5,
      Title: "C Site 2",
      EUMSiteURL: { "Description": "c site2", "Url": "https://envisionitdev.sharepoint.com/sites/csite2" },
      EUMParentURL: { "Description": "EIT Intranet", "Url": "https://envisionitdev.sharepoint.com/sites/eitintranet" }
    }]);
  }
}
