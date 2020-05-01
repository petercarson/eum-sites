import { IDivisionListItem } from "./IDivisionListItem";
import { IChannelTemplateListItem } from "./IChannelTemplateListItem";

export interface ITeamsChannelState {
    hasError: boolean;
    saveSuccess: boolean;
    isLoading: boolean;
    isSaving: boolean;
    errorMessage?: string;
    fieldsValid: boolean;
    channelTemplatesLoaded: boolean;
    channelTemplates: IChannelTemplateListItem[];
    selectedChannelTemplate?: string;
}