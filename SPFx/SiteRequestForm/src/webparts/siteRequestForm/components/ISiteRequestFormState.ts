import { IDivisionListItem } from './IDivisionListItem';
import { ISiteTemplateListItem } from './ISiteTemplateListItem';

export interface ISiteRequestFormState {
    hasError: boolean;
    fieldsValid: boolean;
    saveSuccess: boolean;
    isLoading: boolean;
    isSaving: boolean;
    errorMessage?: string;

    divisionsLoaded: boolean;
    siteTemplatesLoaded: boolean;
    fieldsLoaded: boolean;

    selectedDivision?: string;
    selectedSiteTemplate?: string;
    contentTypeId?: string;
    divisions?: IDivisionListItem[];
    siteTemplates?: ISiteTemplateListItem[];
    fields?: any[];

    alias?: string;
    aliasValidating: boolean;
    aliasIsValid: boolean;

    blacklistedWords?: string[];
    blacklistedWordsLoaded: boolean;

    titleValidating: boolean;
    titleIsValid: boolean;
    
    title?: string;
    
    selectedPrefix: string;
    preselectedPrefix : string;

    externalSharing : any;
    aliasValidationMessage : string;
    siteUrlValidationMessage : string;
}