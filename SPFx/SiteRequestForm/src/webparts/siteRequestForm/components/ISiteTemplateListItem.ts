export interface ISiteTemplateListItem {
    Id: string;
    Title: string;
    TemplateDescription: string;
    ContentTypeName: string;
    UseDefaultForm: boolean;
    Office365Group: boolean;
    SiteVisibilityDefaultValue: string;
    SiteVisibilityShowChoice: boolean;
    CreateTeamDefaultValue: boolean;
    CreateTeamShowToggle: boolean;
    CreateOneNoteShowToggle: boolean;
    CreateOneNoteDefaultValue: boolean;
    CreatePlannerDefaultValue: boolean;
    CreatePlannerShowToggle: boolean; 
    ExternalSharingDefaultValue: string;
    ExternalSharingAllowedOptions: any;
    ExternalSharingShowChoice : boolean;
    DefaultSharingLinkType : string;
    DefaultSharingLinkShowChoice : boolean;
    DefaultLinkPermission : string;
    DefaultLinkPermissionShowChoice : boolean;
}