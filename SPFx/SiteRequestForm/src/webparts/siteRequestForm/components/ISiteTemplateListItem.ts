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
}