define([], function() {
  return {
    "PropertyPaneDescription": "",
    "BasicGroupName": "Site Request Form Settings",
    "DescriptionFieldLabel": "Description",
    "WebPartTitleFieldLabel": "Web Part Title",
    "MasterSiteURLFieldLabel": "Master Site URL",
    "SiteListNameFieldLabel": "Site List Name",
    "DivisionsListNameFieldLabel": "Divisions List Name",
    "SiteTemplatesListNameFieldLabel": "Site Templates List Name",
    "BlacklistedWordsListNameFieldLabel": "Blacklisted Words List Name",
    "DefaultNewItemUrl":"New Item Form URL",

    "ToggleOnText": "Yes",
    "ToggleOffText": "No",

    "PreselectedDivisionLabel": "Preselected Division",
    "PreselectedDivisionDescription" : "Enter the division name as it appears in the Division list to preselect that division in the Division dropdown. The Division dropdown will be hidden.",

    "PreselectedSiteTemplateLabel": "Preselected Site Template",
    "PreselectedSiteTemplateDescription" : "Enter the site template name as it appears in the SiteTemplates list to preselect that template in the Site Template dropdown. The Site Template dropdown will be hidden.",

    "TitleFieldLabel": "Title Field Label",
    "TitleFieldLabelDescription" : "Enter a value here to change the label of the Title field on the form.",

    "DivisionFieldLabel": "Division Field Label",
    "DivisionFieldLabelDescription" : "Enter a value here to change the label of the Division field on the form.",

    "SiteTemplateFieldLabel": "Site Template Field Label",
    "SiteTemplateFieldLabelDescription" : "Enter a value here to change the label of the Site Template field on the form.",

    "SiteProvisioningApiUrlFieldLabel" : "Site provisioning API URL",
    "SiteProvisioningApiUrlFieldDescription": "Enter the site provisioning API URL used to submit site requests. If no URL is provided, webpart will use SharePoint's API to submit - users will require Contribute permissions on the Sites list.",

    "SiteProvisioningApiClientIDFieldLabel": "Site Provisioning API Azure AD Client ID",
    "TenantPropertyDescription": "The web part will retrieve this value from your tenant properties. Enter a value in this field if you wish to override the tenant property.",

    "DivsionDropdownLabel": "Division",
    "DivsionDropdownPlaceholderText": "Select a division",
    "CurrencyFieldSymbol": "$",
    "SiteTemplateDropdownLabel": "Site Template",
    "SiteTemplateDropdownPlaceholderText": "Select a site template",
    
    "DatepickerPlaceholder": "Select a date",
    "LoadingText": "Loading...",
    "SubmitButtonText": "Submit",
    "CancelButtonText": "Cancel",
    "RequiredFieldMessage": "This is a required field.",
    "NumericFieldMessage": "Please enter a numeric value.",
    "CurrencyFieldMessage": "Please enter a currency value (0.00).",
    "InvalidFieldsMessage": "Required fields are missing or invalid.",
    "SaveSuccessMessage": "Your site request has been received. You will be notified when your site is ready.",
    "NoFieldsText": "No fields available",
    "AliasInUseMessage": "This alias is already in use or is invalid.",
    "SiteUrlInUseMessage": "This URL is already in use or is invalid.",
    "BlacklistedWordsMessage": "The following word or phrase cannot be used:"
  }
});