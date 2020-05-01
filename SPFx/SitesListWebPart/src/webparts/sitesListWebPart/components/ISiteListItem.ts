export interface ILookupFieldValue {
    Id: string;
    Title: string;
}
export interface ISiteListItem {
    Id: number;
    Title: string;
    EUMSiteURL: string;
    EUMParentURL: string;
    EUMDivision?: ILookupFieldValue;
    EUMGroupSummary?: string;
    EUMAlias?: string;
    EUMSiteVisibility?: string;
    SitePurpose?: string;
    EUMSiteCreated?: string;
    EUMSiteTemplate?: string;
}