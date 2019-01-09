export interface IUrlFieldValue {
    Description: string;
    Url: string;
}
export interface ISiteListItem {
    Id: number;
    Title: string;
    EUMSiteURL: IUrlFieldValue;
    EUMParentURL: IUrlFieldValue;
}