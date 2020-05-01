declare interface ISiteRequestFormWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;

  WebPartTitleFieldLabel: string;
  MasterSiteURLFieldLabel: string;
  SiteListNameFieldLabel: string;
  DivisionsListNameFieldLabel: string;
  SiteTemplatesListNameFieldLabel: string;
  DefaultNewItemUrl: string;

  ToggleOnText: string;
  ToggleOffText: string;

  TitleFieldLabel: string;
  TitleFieldLabelDescription: string;

  PreselectedDivisionLabel: string;
  PreselectedDivisionDescription: string;

  PreselectedSiteTemplateLabel: string;
  PreselectedSiteTemplateDescription: string;

  SiteProvisioningApiUrlFieldLabel: string;
  SiteProvisioningApiUrlFieldDescription: string;

  DivsionDropdownLabel: string;
  DivsionDropdownPlaceholderText: string;
  SiteTemplateDropdownLabel: string;
  SiteTemplateDropdownPlaceholderText: string;
  DatepickerPlaceholder: string;
  CurrencyFieldSymbol: string;
  SubmitButtonText: string;
  CancelButtonText: string;
  LoadingText: string;

  RequiredFieldMessage: string;
  NumericFieldMessage: string;
  CurrencyFieldMessage: string;
  InvalidFieldsMessage: string;
  SaveSuccessMessage: string;
  NoFieldsText: string;
  AliasInUseMessage: string;
  SiteUrlInUseMessage: string;
}

declare module 'SiteRequestFormWebPartStrings' {
  const strings: ISiteRequestFormWebPartStrings;
  export = strings;
}
