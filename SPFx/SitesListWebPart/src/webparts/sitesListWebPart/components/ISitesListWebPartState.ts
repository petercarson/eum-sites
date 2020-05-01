import { ISiteListItem } from './ISiteListItem';

export interface ISitesListWebPartState {
    hasError: boolean;
    errorMessage?: string;

    sitesLoaded: boolean;

    sitesListItems?: ISiteListItem[];

    currentUserSites?: ISiteListItem[];
    availableSites?: ISiteListItem[];
}