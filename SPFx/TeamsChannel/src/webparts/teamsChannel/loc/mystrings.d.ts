declare interface ITeamsChannelWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;

  WebPartTitleFieldLabel: string;
  MasterSiteUrlFieldLabel: string;
  SitesListNameFieldLabel: string;
  TeamsChannelsListNameFieldLabel: string;
  ChannelTemplatesListNameFieldLabel: string;
  DivisionsListNameFieldLabel: string;

  ToggleOnText: string;
  ToggleOffText: string;

  LoadingText: string;
  SavingText: string;
  SuccessText: string;
  FieldsInvalidErrorText: string;

  InvalidFieldsMessage: string;
  SaveSuccessMessage: string;

  SubmitButtonText: string;

  TeamsChannelTitleFieldLabel: string;
  TeamsChannelTitleFieldPlaceholder: string;
  TeamsChannelPrivacyFieldLabel: string;
  TeamsChannelPrivacyFieldPlaceholder: string;
  TeamsChannelPrivacyFieldWarning: string;
  TeamsChannelPrivacyOptionPrivate: string;
  TeamsChannelPrivacyOptionPublic: string;

  TeamsChannelDescriptionFieldLabel: string;
  TeamsChannelDescriptionFieldPlaceholder: string;

  RequiredFieldMessage: string;

  ChannelTemplateDropdownLabel: string;
  ChannelTemplateDropdownPlaceholderText: string;

  CreateOneNoteSectionToggleLabel: string;
  CreatePlannerToggleLabel: string;
}

declare module 'TeamsChannelWebPartStrings' {
  const strings: ITeamsChannelWebPartStrings;
  export = strings;
}
