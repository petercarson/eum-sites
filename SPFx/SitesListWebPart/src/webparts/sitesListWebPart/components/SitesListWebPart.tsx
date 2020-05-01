import * as React from 'react';
import { Environment, EnvironmentType } from '@microsoft/sp-core-library';
import styles from './SitesListWebPart.module.scss';
import { ISitesListWebPartProps } from './ISitesListWebPartProps';
import { ISitesListWebPartState } from './ISitesListWebPartState';
import { IGroupedSites } from './IGroupedSites';
import * as strings from 'SitesListWebPartWebPartStrings';
import { ISiteListItem } from './ISiteListItem';
import { Pivot, PivotItem, PivotLinkSize } from 'office-ui-fabric-react/lib/Pivot';
import { Link } from 'office-ui-fabric-react/lib/Link';
import { Spinner } from 'office-ui-fabric-react/lib/Spinner';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react/lib/MessageBar';
import { DocumentCard, DocumentCardTitle, IDocumentCardLogoProps, DocumentCardLogo } from 'office-ui-fabric-react/lib/DocumentCard';
import { Fabric } from 'office-ui-fabric-react/lib/Fabric';
import { sp, SearchResults, SearchQuery } from "@pnp/sp";
import { HttpClient, HttpClientResponse } from '@microsoft/sp-http';

export default class SitesListWebPart extends React.Component<ISitesListWebPartProps, ISitesListWebPartState> {
  constructor(props: ISitesListWebPartProps) {
    super(props);

    this.state = {
      hasError: false,
      sitesLoaded: false
    };
  }



  public render(): React.ReactElement<ISitesListWebPartProps> {
    return (
      <Fabric className={styles.sitesListWebPart}>
        {(this.props.webpartTitle) ? <span className={styles.webpartTitle}>{this.props.webpartTitle}</span> : ""}
        {!this.state.sitesLoaded && !this.state.hasError ? this.RenderLoadingSpinner() : ""}
        {(this.state.hasError) ? this.RenderErrors() : ""}

        {!this.state.sitesLoaded && !this.state.hasError ? this.GetSitesListItems() : ""}

        {this.state.sitesLoaded && !this.state.hasError ? this.RenderSitesList() : ""}
      </Fabric>
    );
  }

  private RenderErrors() {
    return (
      <MessageBar messageBarType={MessageBarType.error} isMultiline={true}>{this.state.errorMessage}</MessageBar>
    );
  }

  private RenderNoSitesMessage() {
    return (
      <MessageBar messageBarType={MessageBarType.warning} isMultiline={true}>{strings.NoSitesFoundText}</MessageBar>
    );
  }

  private RenderLoadingSpinner() {
    return (
      <Spinner label={strings.LoadingText} />
    );
  }

  private RenderSitesList() {

    let allSitesDisplay: JSX.Element = this.props.displayAllSites && this.state.sitesListItems && this.state.sitesListItems.length > 0 ? this.GetJsxForSites(this.state.sitesListItems) : null;
    let availableSitesDisplay: JSX.Element = this.props.displayAvailableSites && this.state.availableSites && this.state.availableSites.length > 0 ? this.GetJsxForSites(this.state.availableSites) : null;
    let currentUserSitesDisplay: JSX.Element = this.props.displayUserSites && this.state.currentUserSites && this.state.currentUserSites.length > 0 ? this.GetJsxForSites(this.state.currentUserSites) : null;

    return (
      <Pivot aria-label={this.props.webpartTitle}>

        {currentUserSitesDisplay ? <PivotItem headerText={this.props.tabHeaderUserSites}>
          {currentUserSitesDisplay}
        </PivotItem> : ""}

        {availableSitesDisplay ? <PivotItem headerText={this.props.tabHeaderAvailableSites}>
          {availableSitesDisplay}
        </PivotItem> : ""}

        {allSitesDisplay ? <PivotItem headerText={this.props.tabHeaderAllSites}>
          {allSitesDisplay}
        </PivotItem> : ""}
      </Pivot>
    );
  }


  private GetJsxForSites(sites: ISiteListItem[]) {
    let groupedSites: IGroupedSites[] = this.GroupSites(sites);
    if (groupedSites.length === 0) {
      return this.RenderNoSitesMessage();
    } else if (this.props.displayMode === strings.DisplayModeTabs) {
      return this.RenderPivot(groupedSites);
    } else if (this.props.displayMode === strings.DisplayModeListing) {
      return this.RenderList(groupedSites);
    } else if (this.props.displayMode === strings.DisplayModeTiles) {
      return this.RenderTiles(groupedSites);
    } else {
      return this.RenderAuto(groupedSites);
    }
  }

  private RenderPivot(groupedSites: IGroupedSites[]) {
    return (
      <Fabric>
        <Pivot>
          {groupedSites.map((groupedSite) => {
            return (
              <PivotItem headerText={groupedSite.index}>
                <ul className={styles.sitesList}>
                  {groupedSite.sitesListItems.map((siteListItem) => {
                    return (
                      <li>
                        <Link href={siteListItem.EUMSiteURL} name={siteListItem.Title}>{siteListItem.Title}</Link>
                      </li>
                    );
                  })}
                </ul>
              </PivotItem>
            );
          })};
        </Pivot>
      </Fabric>
    );
  }

  private RenderList(groupedSites: IGroupedSites[]) {
    return (
      <ul className={styles.sitesList}>
        {groupedSites.map((groupedSite) => {
          return (
            groupedSite.sitesListItems.map((siteListItem) => {
              return (
                <li>
                  <Link href={siteListItem.EUMSiteURL} name={siteListItem.Title}>{siteListItem.Title}</Link>
                </li>
              );
            })
          );
        })}
      </ul>
    );
  }

  private RenderTiles(groupedSites: IGroupedSites[]) {
    const logoProps: IDocumentCardLogoProps = {
      logoIcon: 'SharePointLogo'
    };

    return (
      <div className="ms-Grid" dir="ltr">
        <div className="ms-Grid-row">
          {groupedSites.map((groupedSite) => {
            return (
              groupedSite.sitesListItems.map((siteListItem) => {
                return (
                  <div className="ms-Grid-col ms-sm12 ms-md4 ms-lg4 ms-xl3">
                    <DocumentCard className={styles.siteTile} onClickHref={siteListItem.EUMSiteURL}>
                      <DocumentCardLogo {...logoProps} />
                      <DocumentCardTitle
                        title={siteListItem.Title}
                        shouldTruncate={true}
                      />
                      <div className={styles.siteSummary} dangerouslySetInnerHTML={{ __html: siteListItem.EUMGroupSummary }} />
                    </DocumentCard>
                  </div>
                );
              })
            );
          })}
        </div>
      </div>
    );
  }

  private RenderAuto(groupedSites: IGroupedSites[]) {
    if (this.state.sitesListItems.length == 1) {
      window.location.href = this.state.sitesListItems[0].EUMSiteURL;
    } else if (this.state.sitesListItems.length >= 10) {
      return this.RenderPivot(groupedSites);
    } else if (this.state.sitesListItems.length > 1 && this.state.sitesListItems.length < 10) {
      return this.RenderList(groupedSites);
    }
  }

  private GetSitesListItems(): void {
    if (Environment.type === EnvironmentType.Local) {
      this.GetSitesListItemsDebug();
    } else {
      let EUMParentURL: string = null;
      if (this.props.parentSiteURL) {
        // parentSiteURL is set so filter the site list by its children
        EUMParentURL = this.GetParentSiteRelativeUrl();
      } else {
        // no parentSiteURL provided, so assume current site
        EUMParentURL = this.props.currentWebRelativeUrl;
      }

      if (EUMParentURL.toLocaleLowerCase() == this.props.masterSiteURL.toLocaleLowerCase()) {
        this.GetAllSiteListItems();
      } else {
        this.GetSiteListItemsFiltered(EUMParentURL);
      }
    }
  }

  private GetSiteListItemsFiltered(EUMParentURL: string): void {
    if (this.props.displayUserSites) {
      // use SharePoint's search API to retrieve items from the Site Metadata list
      this.GetSiteListMetadata().then((currentUserSites: ISiteListItem[]): void => {
        this.setState({ sitesLoaded: true, currentUserSites: currentUserSites });
        // if we need to display all sites and available sites, retrieve all the sites
        if (this.props.displayAllSites || this.props.displayAvailableSites) {
          this.GetFilteredSites(EUMParentURL).then((allSites: ISiteListItem[]): void => {
            let availableSites: ISiteListItem[] = [];
            if (this.props.displayAvailableSites) {
              // diff the list of current user sites and available sites
              availableSites = this.GetAvailableSites(currentUserSites, allSites);
            }
            this.setState({ sitesLoaded: true, sitesListItems: allSites, availableSites: availableSites });
          }, (e: Error): void => {
            this.setState({ sitesLoaded: false, hasError: true, errorMessage: e.message });
          });
        }
      }, (e: Error): void => {
        this.setState({ sitesLoaded: false, hasError: true, errorMessage: e.message });
      });
    } else if (!this.props.displayUserSites && this.props.displayAvailableSites || this.props.displayAllSites) {
      // get all sites
      this.GetFilteredSites(EUMParentURL).then((allSites: ISiteListItem[]): void => {
        this.setState({ sitesLoaded: true, sitesListItems: allSites });
      }, (e: Error): void => {
        this.setState({ sitesLoaded: false, hasError: true, errorMessage: e.message });
      });
    }
  }

  private async GetFilteredSites(EUMParentURL: string): Promise<ISiteListItem[]> {
    if (this.props.siteProvisioningApiUrl) {
      // if the API endpoint was provided, then call the API to retrieve all sites
      return await this.GetFilteredSitesFromApi(EUMParentURL);
    } else {
      // directly query the Sites list
      // get all children of the parent url
      let EUMParentURLAbsolute: string = `${this.GetRootSiteUrl()}${EUMParentURL}`;
      let viewXml: string = `<View>
        <Query>
          <Where>
            <And>
              <And>
                <Or>
                  <Eq><FieldRef Name="EUMParentURL"/><Value Type="Text">${EUMParentURL}</Value></Eq>
                  <Eq><FieldRef Name="EUMParentURL"/><Value Type="Text">${EUMParentURLAbsolute}</Value></Eq>
                </Or>
                <Neq><FieldRef Name='EUMSiteVisibility'/><Value Type='Text'>Hidden</Value></Neq>
              </And>
              <IsNotNull>
                <FieldRef Name='EUMSiteCreated'/>
              </IsNotNull>
            </And>
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
          <FieldRef Name="EUMDivision"></FieldRef>
          <FieldRef Name="EUMGroupSummary"></FieldRef>
        </ViewFields>
      </View>`;

      return sp.web
        .lists
        .getByTitle(this.props.siteListName)
        .getItemsByCAMLQuery({ ViewXml: viewXml });
    }
  }

  private GetAllSiteListItems(): void {
    if (this.props.displayUserSites) {
      // use SharePoint's search API to retrieve items from the Site Metadata list
      this.GetSiteListMetadata().then((currentUserSites: ISiteListItem[]): void => {
        this.setState({ sitesLoaded: true, currentUserSites: currentUserSites });
        // if we need to display all sites and available sites, retrieve all the sites
        if (this.props.displayAllSites || this.props.displayAvailableSites) {
          this.GetAllSites().then((allSites: ISiteListItem[]): void => {
            let availableSites: ISiteListItem[] = [];
            if (this.props.displayAvailableSites) {
              // diff the list of current user sites and available sites
              availableSites = this.GetAvailableSites(currentUserSites, allSites);
            }
            this.setState({ sitesLoaded: true, sitesListItems: allSites, availableSites: availableSites });
          }, (e: Error): void => {
            this.setState({ sitesLoaded: false, hasError: true, errorMessage: e.message });
          });
        }
      }, (e: Error): void => {
        this.setState({ sitesLoaded: false, hasError: true, errorMessage: e.message });
      });
    } else if (!this.props.displayUserSites && this.props.displayAvailableSites || this.props.displayAllSites) {
      // get all sites
      this.GetAllSites().then((allSites: ISiteListItem[]): void => {
        this.setState({ sitesLoaded: true, sitesListItems: allSites });
      }, (e: Error): void => {
        this.setState({ sitesLoaded: false, hasError: true, errorMessage: e.message });
      });
    }
  }

  private async GetAllSites(): Promise<ISiteListItem[]> {
    if (this.props.siteProvisioningApiUrl) {
      // if the API endpoint was provided, then call the API to retrieve all sites
      return await this.GetAllSitesFromApi();
    } else {
      // directly query the Sites list
      return sp.web
        .lists
        .getByTitle(this.props.siteListName)
        .items
        .select("ID", "Title", "EUMSiteURL", "EUMParentURL", "EUMDivision/ID", "EUMDivision/Title", "EUMGroupSummary")
        .expand("EUMDivision")
        .top(5000)
        .filter("EUMSiteCreated ne null and EUMSiteVisibility ne 'Hidden'")
        .orderBy("EUMDivision/Title", true)
        .orderBy("Title", true)
        .get();
    }
  }

  private async GetAllSitesFromApi(): Promise<ISiteListItem[]> {
    let httpResponse = await this.props.HttpClient.get(`${this.props.siteProvisioningApiUrl}/Sites`, HttpClient.configurations.v1, {
      headers: {
        'authorization': `Bearer ${this.props.accessToken}`
      }
    })
      .then((response: HttpClientResponse): Promise<ISiteListItem[]> => {
        return new Promise<ISiteListItem[]>((resolve, reject) => {
          if (response.ok && response.status == 200) {
            response.json().then((responseJson) => {
              resolve(responseJson);
            });
          } else {
            response.text().then((responseText) => {
              reject(responseText);
            });
          }
        });
      });

    return httpResponse;
  }

  private async GetFilteredSitesFromApi(EUMParentURL: string): Promise<ISiteListItem[]> {
    let httpResponse = await this.props.HttpClient.get(`${this.props.siteProvisioningApiUrl}/Sites?parentUrl=${EUMParentURL}`, HttpClient.configurations.v1, {
      headers: {
        'authorization': `Bearer ${this.props.accessToken}`
      }
    })
      .then((response: HttpClientResponse): Promise<ISiteListItem[]> => {
        return new Promise<ISiteListItem[]>((resolve, reject) => {
          if (response.ok && response.status == 200) {
            response.json().then((responseJson) => {
              resolve(responseJson);
            });
          } else {
            response.text().then((responseText) => {
              reject(responseText);
            });
          }
        });
      });

    return httpResponse;
  }

  private async GetSiteListMetadata(): Promise<ISiteListItem[]> {
    let results: SearchResults = await sp.search({
      Querytext: this.props.siteMetadataSearchQuery,
      SelectProperties: this.props.siteMetadataManagedProperties ? this.props.siteMetadataManagedProperties.split(',') : [],
      RowLimit: 500
    });

    let sites: ISiteListItem[] = [];
    if (results.PrimarySearchResults.length > 0) {
      results.PrimarySearchResults.forEach((result) => {
        let site: ISiteListItem = {
          Id: result['ListItemId'],
          Title: result['Title'],
          EUMSiteURL: result['EUMSiteURL'],
          EUMParentURL: result['EUMParentURL'],
          EUMDivision: {
            Id: "0",
            Title: result['EUMDivision']
          },
          EUMGroupSummary: result['EUMGroupSummary'],
          EUMAlias: result['EUMAlias'],
          EUMSiteVisibility: result['EUMSiteVisibility'],
          SitePurpose: result['SitePurpose'],
          EUMSiteCreated: result['EUMSiteCreated'],
          EUMSiteTemplate: result['EUMSiteTemplate']
        };

        sites.push(site);
      });
    }

    return sites.sort(function (a, b) {
      return a.Title === b.Title ? 0 : a.Title < b.Title ? -1 : 1;
    });
  }


  private GetAvailableSites(currentUserSites: ISiteListItem[], allSites: ISiteListItem[]): ISiteListItem[] {
    let availableSites: ISiteListItem[] = [];

    // return all items in allSites that are not in currentUserSites
    availableSites = allSites.filter(s => !(currentUserSites.some(c => c.EUMSiteURL.toLocaleLowerCase() == s.EUMSiteURL.toLocaleLowerCase())));

    return availableSites;
  }

  private GetSitesListItemsDebug(): void {
    this.GetMockItems()
      .then((items: ISiteListItem[]): void => {
        this.setState({ sitesLoaded: true, sitesListItems: items });
      });
  }

  private GetMockItems(): Promise<ISiteListItem[]> {
    return Promise.resolve([{
      Id: 0,
      Title: "A Site",
      EUMSiteURL: "",
      EUMParentURL: "https://envisionitdev.sharepoint.com/sites/eitintranet",
      EUMGroupSummary: "A Site for testing purposes"
    },
    {
      Id: 1,
      Title: "B Site",
      EUMSiteURL: "https://envisionitdev.sharepoint.com/sites/bsite",
      EUMParentURL: "https://envisionitdev.sharepoint.com/sites/eitintranet",
      EUMGroupSummary: "A Site for testing purposes"
    },
    {
      Id: 3,
      Title: "C Site",
      EUMSiteURL: "https://envisionitdev.sharepoint.com/sites/csite",
      EUMParentURL: "https://envisionitdev.sharepoint.com/sites/eitintranet",
      EUMGroupSummary: "A Site for testing purposes"
    },
    {
      Id: 4,
      Title: "1 Site",
      EUMSiteURL: "https://envisionitdev.sharepoint.com/sites/numsite",
      EUMParentURL: "https://envisionitdev.sharepoint.com/sites/eitintranet",
      EUMGroupSummary: "A Site for testing purposes"
    }, {
      Id: 5,
      Title: "C Site 2",
      EUMSiteURL: "https://envisionitdev.sharepoint.com/sites/csite2",
      EUMParentURL: "https://envisionitdev.sharepoint.com/sites/eitintranet",
      EUMGroupSummary: "A Site for testing purposes"
    }]);
  }

  private GroupSites(sites: ISiteListItem[]): IGroupedSites[] {
    if (this.props.groupBy === strings.GroupByParent) {
      return this.GroupByDivision(sites);
    } else {
      return this.GroupSitesByTitle(sites);
    }
  }

  private GroupSitesByTitle(sites: ISiteListItem[]): IGroupedSites[] {
    let groupedSites: IGroupedSites[] = [];
    sites.forEach((currentSite) => {
      let firstLetter: any = null;
      if (currentSite && currentSite.Title && currentSite.EUMSiteURL && currentSite.EUMSiteURL) {
        firstLetter = (currentSite.Title.charAt(0).toUpperCase());
        if (!isNaN(firstLetter)) {
          firstLetter = "#";
        }

        let updateRequired: boolean = groupedSites.some(s => s.index === firstLetter);
        if (updateRequired) {
          groupedSites = groupedSites.map(s => {
            if (s.index === firstLetter) {
              s.sitesListItems.push(currentSite);
              return s;
            } else {
              return s;
            }
          });
        } else {
          let groupedSite: IGroupedSites = { index: firstLetter, sitesListItems: [] };
          groupedSite.sitesListItems.push(currentSite);
          groupedSites.push(groupedSite);
        }
      }
    });

    return groupedSites;
  }

  private GroupByDivision(sites: ISiteListItem[]): IGroupedSites[] {
    let groupedSites: IGroupedSites[] = [];
    sites.forEach((currentSite) => {
      let divisionName: string = null;
      if (currentSite && currentSite.Title && currentSite.EUMSiteURL && currentSite.EUMSiteURL) {
        divisionName = currentSite.EUMDivision ? currentSite.EUMDivision.Title : null;
        if (divisionName) {
          let updateRequired: boolean = groupedSites.some(s => s.index === divisionName);
          if (updateRequired) {
            groupedSites = groupedSites.map(s => {
              if (s.index === divisionName) {
                s.sitesListItems.push(currentSite);
                return s;
              } else {
                return s;
              }
            });
          } else {
            let groupedSite: IGroupedSites = { index: divisionName, sitesListItems: [] };
            groupedSite.sitesListItems.push(currentSite);
            groupedSites.push(groupedSite);
          }
        }
      }
    });

    return groupedSites;
  }

  // *********************
  // Url Parsing Helpers
  // *********************
  private GetRootSiteUrl(): string {
    return (new URL(this.props.currentWebAbsoluteUrl)).origin;
  }

  private GetParentSiteRelativeUrl(): string {
    return this.props.parentSiteURL.replace(this.GetRootSiteUrl(), "");
  }
}
