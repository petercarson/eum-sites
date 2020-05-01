import * as React from 'react';
import { ISiteRequestFormProps } from './ISiteRequestFormProps';
import { ISiteRequestFormState } from './ISiteRequestFormState';
import { IFieldValue } from './IFieldValue';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react/lib/MessageBar';
import { Toggle } from 'office-ui-fabric-react/lib/Toggle';
import { Dropdown } from 'office-ui-fabric-react/lib/Dropdown';
import { DatePicker } from 'office-ui-fabric-react/lib/DatePicker';
import { PrimaryButton, DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { Spinner } from 'office-ui-fabric-react/lib/Spinner';
import styles from './SiteRequestForm.module.scss';
import { IDivisionListItem } from './IDivisionListItem';
import { ISiteTemplateListItem } from './ISiteTemplateListItem';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { IPersonaProps } from 'office-ui-fabric-react/lib/components/Persona/Persona.types';
import { TaxonomyPicker, IPickerTerms } from "@pnp/spfx-controls-react/lib/TaxonomyPicker";
import { ListItemPicker } from '@pnp/spfx-controls-react/lib/listItemPicker';
import { RichText } from "@pnp/spfx-controls-react/lib/RichText";
import { Label } from 'office-ui-fabric-react/lib/Label';
import * as strings from 'SiteRequestFormWebPartStrings';
import { sp, ItemAddResult, Web, SPHttpClient } from "@pnp/sp";
import { HttpClient, HttpClientResponse } from '@microsoft/sp-http';
import { IBlacklistedWordsListItem } from './IBlacklistedWordsListItem';

export default class SiteRequestForm extends React.Component<ISiteRequestFormProps, ISiteRequestFormState> {
  private ModernTeamSite: boolean = false;

  private PreselectedDivision: IDivisionListItem = null;
  private PreselectedSiteTemplate: ISiteTemplateListItem = null;

  private SiteVisibilityDefault: string = 'Public';
  private ShowSiteVisibilityDropdown: boolean = true;
  private CreateTeamDefault: boolean = true;
  private ShowCreateTeamToggle: boolean = true;
  private CreateOneNoteDefault: boolean = true;
  private ShowCreateOneNoteToggle: boolean = true;
  private CreatePlannerDefault: boolean = true;
  private ShowCreatePlannerToggle: boolean = true;

  constructor(props: ISiteRequestFormProps) {
    super(props);

    this.state = {
      saveSuccess: false,
      hasError: false,
      fieldsValid: true,
      isLoading: true,
      isSaving: false,
      fieldsLoaded: false,
      divisionsLoaded: false,
      blacklistedWordsLoaded: false,
      siteTemplatesLoaded: false,
      alias: null,
      aliasValidating: false,
      aliasIsValid: false,
      titleValidating: false,
      titleIsValid: false,
      title: null,
      selectedPrefix: ''
    };
  }

  private fieldValues: IFieldValue[] = [];

  public render() {
    let getBlacklistedWords: boolean = (!this.state.hasError && !this.state.blacklistedWordsLoaded);
    let getDivisions: boolean = (!this.state.hasError && !this.state.divisionsLoaded && this.state.blacklistedWordsLoaded);
    let renderDivisionsDropdown: boolean = this.state.divisionsLoaded;
    let getSiteTemplates: boolean = (!this.state.hasError && !this.state.siteTemplatesLoaded && this.state.selectedDivision != null);
    let renderSiteTemplatesDropdown: boolean = (this.state.siteTemplatesLoaded && this.state.selectedDivision != null);
    let getContentTypeFields: boolean = (!this.state.hasError && !this.state.fieldsLoaded && this.state.selectedSiteTemplate != null);
    let renderFields: boolean = (this.state.fieldsLoaded && this.state.selectedSiteTemplate != null);

    return (
      <div id="SiteRequestForm" className={styles.siteRequestForm}>
        {(this.props.webpartTitle) ? <span className={styles.webpartTitle}>{this.props.webpartTitle}</span> : ""}
        {(this.state.saveSuccess) ? this.RenderSuccess() : ""}
        <div hidden={this.state.saveSuccess}>
          {(this.props.description) ? <p>{this.props.description}</p> : ""}

          {getBlacklistedWords ? this.GetBlacklistedWords() : ""}
          {getDivisions ? this.GetDivisions() : ""}
          {renderDivisionsDropdown ? this.RenderDivisionsDropdown() : ""}

          {getSiteTemplates ? this.GetSiteTemplates() : ""}
          {renderSiteTemplatesDropdown ? this.RenderSiteTemplatesDropdown() : ""}

          <div hidden={!this.state.divisionsLoaded || !this.state.siteTemplatesLoaded}>
            {getContentTypeFields ? this.GetContentTypeFields() : ""}
            {renderFields ? this.RenderFields() : ""}
          </div>

          {(this.state.isLoading || this.state.isSaving) ? this.RenderLoadingSpinner() : ""}
          {(this.state.hasError) ? this.RenderErrors() : ""}
          {(!this.state.fieldsValid) ? this.RenderInvalidFieldsMessage() : ""}
        </div>
      </div>
    );
  }


  private RenderErrors() {
    return (
      <MessageBar messageBarType={MessageBarType.error} isMultiline={true}>{this.state.errorMessage}</MessageBar>
    );
  }

  private RenderInvalidFieldsMessage() {
    return (
      <MessageBar messageBarType={MessageBarType.warning} isMultiline={true}>{strings.InvalidFieldsMessage}</MessageBar>
    );
  }

  private RenderSuccess() {
    return (
      <MessageBar messageBarType={MessageBarType.success} isMultiline={true}>{strings.SaveSuccessMessage}</MessageBar>
    );
  }

  private RenderLoadingSpinner() {
    return (
      <Spinner label={strings.LoadingText} />
    );
  }

  private GetBlacklistedWords(): void {
    sp.web.lists
      .getByTitle(this.props.blacklistedWordsListName)
      .items
      .top(1) // there should only be one entry at most
      .get()
      .then((items: IBlacklistedWordsListItem[]): void => {
        let blacklistedWords: string[] = [];

        if (items.length == 1 && items[0].BlacklistedWordsCSV) {
          blacklistedWords = items[0].BlacklistedWordsCSV.split(',');
        }

        this.setState({ blacklistedWords: blacklistedWords, blacklistedWordsLoaded: true });
      }).catch((e: Error): void => {
        this.setState({ isLoading: false, hasError: true, errorMessage: e.message });
      });
  }

  private GetDivisions(): void {
    sp.web.lists
      .getByTitle(this.props.divisionsListName)
      .items
      .get()
      .then((items: IDivisionListItem[]): void => {
        this.PreselectedDivision = null;
        // if only 1 item returned, then default to that item
        if (items.length === 1) {
          this.setState({ isLoading: false, divisions: items, divisionsLoaded: true, selectedDivision: items[0].Id, selectedPrefix: items[0].Prefix });
        } else if (this.props.preselectedDivision) {
          // if preselected division specified, retrieve that division from the list of items returned and default to that item
          let filteredItems: IDivisionListItem[] = items.filter(d => d.Title === this.props.preselectedDivision);
          if (filteredItems.length > 0) {
            this.PreselectedDivision = filteredItems[0];
            this.setState({ isLoading: false, divisions: items, divisionsLoaded: true, selectedDivision: this.PreselectedDivision.Id, selectedPrefix: this.PreselectedDivision.Prefix });
          } else {
            this.setState({ isLoading: false, divisions: items, divisionsLoaded: true });
          }
        } else {
          this.setState({ isLoading: false, divisions: items, divisionsLoaded: true });
        }
      }).catch((e: Error): void => {
        this.setState({ isLoading: false, hasError: true, errorMessage: e.message });
      });

  }

  private RenderDivisionsDropdown() {
    if (this.state.divisions.length <= 1 || this.PreselectedDivision) {
      return null;
    }

    return (
      <div id="DivisionDropdown">
        <Dropdown
          title="EUMDivisionId"
          id="EUMDivisionId"
          label={this.props.divisionFieldLabel ? this.props.divisionFieldLabel : strings.DivsionDropdownLabel}
          placeholder={strings.DivsionDropdownPlaceholderText}
          required={true}
          options={this.state.divisions.map(division => ({ key: division.Id, text: division.Title, data: { prefix: division.Prefix } }))}
          onChanged={(item) => { this.setState({ isLoading: true, siteTemplatesLoaded: false, selectedSiteTemplate: null, selectedDivision: item.key.toString(), selectedPrefix: item.data.prefix }); }}
          disabled={this.state.isSaving || this.state.isLoading}
        />
      </div>
    );
  }

  private GetSiteTemplates(): void {
    let camlViewXml: string = `<View><Query><Where><Eq><FieldRef Name='Divisions' LookupId='TRUE' /><Value Type='Lookup'>${this.state.selectedDivision}</Value></Eq></Where></Query></View>`;

    sp.web.lists
      .getByTitle(this.props.siteTemplatesListName)
      .getItemsByCAMLQuery({ ViewXml: camlViewXml })
      .then((items: ISiteTemplateListItem[]): void => {
        this.PreselectedSiteTemplate = null;

        // if only 1 item returned, then default to that item
        if (items.length === 1) {
          this.setState({ isLoading: false, fieldsLoaded: false, siteTemplates: items, siteTemplatesLoaded: true, selectedSiteTemplate: items[0].Id });
        } else if (this.props.preselectedSiteTemplate) {
          // if preselected template specified, retrieve that template from the list of items returned and default to that item
          let filteredItems: ISiteTemplateListItem[] = items.filter(d => d.Title === this.props.preselectedSiteTemplate);
          if (filteredItems.length > 0) {
            this.PreselectedSiteTemplate = filteredItems[0];
            this.setState({ isLoading: false, fieldsLoaded: false, siteTemplates: items, siteTemplatesLoaded: true, selectedSiteTemplate: this.PreselectedSiteTemplate.Id });
          } else {
            this.setState({ isLoading: false, fieldsLoaded: false, siteTemplates: items, siteTemplatesLoaded: true });
          }
        } else {
          this.setState({ isLoading: false, fieldsLoaded: false, siteTemplates: items, siteTemplatesLoaded: true });
        }
      }).catch((e: Error): void => {
        this.setState({ isLoading: false, hasError: true, errorMessage: e.message });
      });
  }


  private RenderSiteTemplatesDropdown() {
    if (this.state.siteTemplates.length <= 1 || this.PreselectedSiteTemplate) {
      return null;
    }
    return (
      <div id="SiteTemplateDropdownSection">
        <Dropdown
          title="EUMSiteTemplate"
          id="EUMSiteTemplate"
          label={this.props.siteTemplateFieldLabel ? this.props.siteTemplateFieldLabel : strings.SiteTemplateDropdownLabel}
          placeholder={strings.SiteTemplateDropdownPlaceholderText}
          required={true}
          options={this.state.siteTemplates.map(siteTemplate => ({ key: siteTemplate.Id, text: siteTemplate.Title }))}
          onChanged={(item) => this.setState({ isLoading: true, fieldsLoaded: false, selectedSiteTemplate: item.key.toString(), alias: '' })}
          disabled={this.state.isSaving || this.state.isLoading}
        />
      </div>
    );
  }

  private GetContentTypeFields(): void {
    let contentTypeName: string = this.state.siteTemplates.filter(s => s.Id == this.state.selectedSiteTemplate)[0].ContentTypeName;
    let office365Group: boolean = this.state.siteTemplates.filter(s => s.Id == this.state.selectedSiteTemplate)[0].Office365Group;

    this.SiteVisibilityDefault = this.state.siteTemplates.filter(s => s.Id == this.state.selectedSiteTemplate)[0].SiteVisibilityDefaultValue;
    this.ShowSiteVisibilityDropdown = this.state.siteTemplates.filter(s => s.Id == this.state.selectedSiteTemplate)[0].SiteVisibilityShowChoice;

    this.CreateTeamDefault = this.state.siteTemplates.filter(s => s.Id == this.state.selectedSiteTemplate)[0].CreateTeamDefaultValue;
    this.ShowCreateTeamToggle = this.state.siteTemplates.filter(s => s.Id == this.state.selectedSiteTemplate)[0].CreateTeamShowToggle;

    this.CreateOneNoteDefault = this.state.siteTemplates.filter(s => s.Id == this.state.selectedSiteTemplate)[0].CreateOneNoteDefaultValue;
    this.ShowCreateOneNoteToggle = this.state.siteTemplates.filter(s => s.Id == this.state.selectedSiteTemplate)[0].CreateOneNoteShowToggle;

    this.CreatePlannerDefault = this.state.siteTemplates.filter(s => s.Id == this.state.selectedSiteTemplate)[0].CreatePlannerDefaultValue;
    this.ShowCreatePlannerToggle = this.state.siteTemplates.filter(s => s.Id == this.state.selectedSiteTemplate)[0].CreatePlannerShowToggle;

    this.ModernTeamSite = false;
    sp.web.contentTypes
      .filter(`Name eq '${contentTypeName}'`)
      .expand("Fields")
      .get()
      .then((data: any[]): void => {
        if (data && data.length > 0) {
          let fields: any[] = data[0].Fields.results.filter(f => !f.Title.startsWith("Hide_Form_"));

          // if office365Group, then show Alias, Create Team, and hide Site URL. Otherwise, hide Alias and show site URL
          if (office365Group) {
            fields = fields.filter(f => !(f.InternalName === "EUMSiteURL"));
            this.ModernTeamSite = true;
          } else {
            fields = fields.filter(f => !(f.InternalName === "EUMAlias") &&
              !(f.InternalName === "EUMCreateTeam") &&
              !(f.InternalName === "EUMCreateOneNote") &&
              !(f.InternalName === "EUMCreatePlanner"));
          }

          // hide the toggles if specified
          if (!this.ShowSiteVisibilityDropdown) {
            fields = fields.filter(f => !(f.InternalName === "EUMSiteVisibility"));
          }

          if (this.ModernTeamSite && !this.ShowCreateTeamToggle) {
            fields = fields.filter(f => !(f.InternalName === "EUMCreateTeam"));
          }

          if (this.ModernTeamSite && !this.ShowCreateOneNoteToggle) {
            fields = fields.filter(f => !(f.InternalName === "EUMCreateOneNote"));
          }

          if (this.ModernTeamSite && !this.ShowCreatePlannerToggle) {
            fields = fields.filter(f => !(f.InternalName === "EUMCreatePlanner"));
          }

          let contentTypeId: string = data[0].Id.StringValue;
          this.setState({ isLoading: false, fields: fields, fieldsLoaded: true, contentTypeId: contentTypeId });
        }
      }).catch((e: Error): void => {
        this.setState({ isLoading: false, hasError: true, errorMessage: e.message });
      });
  }

  private RenderFields() {
    return (this.state.fields && (this.state.fields.length > 0)) ?
      this.state.fields.map((spField) => {

        if (!spField.Hidden) {
          if (spField.InternalName === 'Title') {
            return (
              <div>
                <TextField
                  id={spField.InternalName}
                  name={spField.InternalName}
                  label={this.props.titleFieldLabel ? this.props.titleFieldLabel : spField.Title}
                  multiline={spField.TypeAsString === 'Note'}
                  required={true}
                  onChanged={(value) => { this.SaveTitleFieldValue(spField.InternalName, value, spField.TypeAsString); }}
                  validateOnLoad={false}
                  onGetErrorMessage={(value) => this.ValidateTitleField(value, true)}
                  disabled={this.state.isSaving || this.state.isLoading}
                  prefix={this.state.selectedPrefix}
                  value= {this.state.title}
                />
                <Label htmlFor={spField.InternalName}>{spField.Description}</Label>
              </div>
            );
          }

          if (spField.InternalName === 'EUMAlias') {
            return (
              <div>
                <TextField
                  id={spField.InternalName}
                  name={spField.InternalName}
                  label={spField.Title}
                  multiline={spField.TypeAsString === 'Note'}
                  required={true}
                  onChanged={(value) => this.SaveAliasFieldValue(value)}
                  validateOnLoad={false}
                  onGetErrorMessage={(value) => this.ValidateAliasField(value, true)}
                  disabled={this.state.isSaving || this.state.isLoading}
                  prefix={this.state.selectedPrefix}
                  value={this.state.alias}
                />
                <Label htmlFor={spField.InternalName}>{spField.Description}</Label>
              </div>
            );
          }

          if (spField.InternalName === 'EUMSiteURL') {
            return (
              <div>
                <TextField
                  id={spField.InternalName}
                  name={spField.InternalName}
                  label={spField.Title}
                  multiline={spField.TypeAsString === 'Note'}
                  required={true}
                  onChanged={(value) => this.SaveAliasFieldValue(value)}
                  validateOnLoad={false}
                  onGetErrorMessage={(value) => this.ValidateSiteUrlField(value, true)}
                  disabled={this.state.isSaving || this.state.isLoading}
                  prefix={this.state.selectedPrefix ? `/sites/${this.state.selectedPrefix}` : '/sites/'}
                  value={this.state.alias}
                />
                <Label htmlFor={spField.InternalName}>{spField.Description}</Label>
              </div>
            );
          }

          if (spField.InternalName === 'EUMSiteVisibility') {
            let options: string[] = spField.Required ? spField.Choices.results : [''].concat(spField.Choices.results);
            return (
              <div>
                <Dropdown
                  id={spField.InternalName}
                  title={spField.InternalName}
                  label={spField.Title}
                  required={true}
                  multiSelect={false}
                  options={options.map((option: string) => ({ key: option, text: option }))}
                  onChanged={(value) => this.SaveFieldValue(spField.InternalName, value.key.toString(), spField.TypeAsString)}
                  disabled={this.state.isSaving || this.state.isLoading}
                  defaultSelectedKey={this.SiteVisibilityDefault}
                />
                <Label htmlFor={spField.InternalName}>{spField.Description}</Label>
              </div>
            );
          }

          if ((!(spField.InternalName === 'Title') && !(spField.InternalName === 'EUMAlias') && !(spField.InternalName === 'EUMAlias')) &&
            spField.TypeAsString === 'Text' || spField.TypeAsString === 'Note') {
            return (
              <div>
                <TextField
                  id={spField.InternalName}
                  name={spField.InternalName}
                  label={spField.Title}
                  multiline={spField.TypeAsString === 'Note'}
                  required={spField.Required}
                  onChanged={(value) => this.SaveFieldValue(spField.InternalName, value, spField.TypeAsString)}
                  validateOnLoad={false}
                  onGetErrorMessage={(value) => this.ValidateRequiredField(value, spField.Required)}
                  disabled={this.state.isSaving || this.state.isLoading}
                />
                <Label htmlFor={spField.InternalName}>{spField.Description}</Label>
              </div>
            );
          }

          if (spField.TypeAsString === 'HTML') {
            return (
              <div id={spField.InternalName}>
                <Label required={spField.Required}>{spField.Title}</Label>
                <div style={{ position: 'relative' }}>
                  <RichText
                    className={styles.richText}
                    onChange={(value) => this.SaveHtmlFieldValue(spField.InternalName, value)}
                  />
                </div>
              </div>
            );
          }

          if (spField.TypeAsString === 'URL') {
            return (
              <div>
                <TextField
                  id={spField.InternalName}
                  name={spField.InternalName}
                  label={spField.Title}
                  required={spField.Required}
                  onChanged={(value) => this.SaveUrlFieldValue(spField.InternalName, value, spField.TypeAsString)}
                  validateOnLoad={false}
                  onGetErrorMessage={(value) => this.ValidateRequiredField(value, spField.Required)}
                  disabled={this.state.isSaving || this.state.isLoading}
                />
                <Label htmlFor={spField.InternalName}>{spField.Description}</Label>
              </div>
            );
          }

          if (spField.TypeAsString === 'Boolean') {
            return (
              <div>
                <Toggle
                  id={spField.InternalName}
                  defaultChecked={(spField.InternalName !== 'EUMCreateTeam' && spField.InternalName !== 'EUMCreateOneNote' && spField.InternalName !== 'EUMCreatePlanner' && spField.DefaultValue === '1') ||
                    (spField.InternalName === 'EUMCreateTeam' && this.CreateTeamDefault) ||
                    (spField.InternalName === 'EUMCreateOneNote' && this.CreateOneNoteDefault) ||
                    (spField.InternalName === 'EUMCreatePlanner' && this.CreatePlannerDefault)
                  }
                  label={spField.Title}
                  onText={strings.ToggleOnText}
                  offText={strings.ToggleOffText}
                  onChanged={(value) => this.SaveFieldValue(spField.InternalName, value, spField.TypeAsString)}
                  disabled={this.state.isSaving || this.state.isLoading}
                />
                <Label htmlFor={spField.InternalName}>{spField.Description}</Label>
              </div>
            );
          }

          if (spField.TypeAsString === 'Choice') {
            let options: string[] = spField.Required ? spField.Choices.results : [''].concat(spField.Choices.results);
            return (
              <div>
                <Dropdown
                  id={spField.InternalName}
                  title={spField.InternalName}
                  label={spField.Title}
                  required={spField.Required}
                  multiSelect={false}
                  options={options.map((option: string) => ({ key: option, text: option }))}
                  onChanged={(value) => this.SaveFieldValue(spField.InternalName, value.key.toString(), spField.TypeAsString)}
                  disabled={this.state.isSaving || this.state.isLoading}
                />
                <Label htmlFor={spField.InternalName}>{spField.Description}</Label>
              </div>
            );
          }

          if (spField.TypeAsString === 'MultiChoice') {
            let options: string[] = spField.Choices.results;
            return (
              <div>
                <Dropdown
                  id={spField.InternalName}
                  title={spField.InternalName}
                  label={spField.Title}
                  required={spField.Required}
                  multiSelect={true}
                  options={options.map((option: string) => ({ key: option, text: option }))}
                  onChanged={(value) => this.SaveMultiSelectValue(spField.InternalName, value)}
                  disabled={this.state.isSaving || this.state.isLoading}
                />
                <Label htmlFor={spField.InternalName}>{spField.Description}</Label>
              </div>
            );
          }

          if (spField.TypeAsString === 'Number') {
            return (
              <div>
                <TextField
                  id={spField.InternalName}
                  name={spField.InternalName}
                  label={spField.Title}
                  required={spField.Required}
                  onChanged={(value) => this.SaveFieldValue(spField.InternalName, value, spField.TypeAsString)}
                  onGetErrorMessage={(value) => this.ValidateNumericField(value, spField.Required, spField.TypeAsString)}
                  disabled={this.state.isSaving || this.state.isLoading}
                />
                <Label htmlFor={spField.InternalName}>{spField.Description}</Label>
              </div>
            );
          }

          if (spField.TypeAsString === 'Currency') {
            return (
              <div>
                <TextField
                  id={spField.InternalName}
                  name={spField.InternalName}
                  label={`${spField.Title} (${strings.CurrencyFieldSymbol})`}
                  required={spField.Required}
                  onChanged={(value) => this.SaveFieldValue(spField.InternalName, value, spField.TypeAsString)}
                  onGetErrorMessage={(value) => this.ValidateNumericField(value, spField.Required, spField.TypeAsString)}
                  disabled={this.state.isSaving || this.state.isLoading}
                />
                <Label htmlFor={spField.InternalName}>{spField.Description}</Label>
              </div>
            );
          }

          if (spField.TypeAsString === 'DateTime') {
            return (
              <div id={spField.InternalName}>
                <DatePicker
                  label={spField.Title}
                  isRequired={spField.Required}
                  firstWeekOfYear={1}
                  showMonthPickerAsOverlay={true}
                  allowTextInput={true}
                  placeholder={strings.DatepickerPlaceholder}
                  ariaLabel={strings.DatepickerPlaceholder}
                  formatDate={(date: Date) => date.getDate() + '/' + (date.getMonth() + 1) + '/' + (date.getFullYear())}
                  parseDateFromString={(dateStr: string) => new Date(Date.parse(dateStr))}
                  onSelectDate={(value) => this.SaveDateFieldValue(spField.InternalName, value)}
                  disabled={this.state.isSaving || this.state.isLoading}
                />
                <Label htmlFor={spField.InternalName}>{spField.Description}</Label>
              </div>
            );
          }

          if (spField.TypeAsString === 'User') {
            return (
              <div id={spField.InternalName}>
                <PeoplePicker
                  context={this.props.context}
                  titleText={spField.Title}
                  isRequired={spField.Required}
                  ensureUser={true}
                  selectedItems={(items) => { this.SavePeoplePickerValue(`${spField.InternalName}`, items); }}
                  principalTypes={spField.SelectionMode === 1 ?
                    [PrincipalType.User, PrincipalType.SecurityGroup, PrincipalType.SharePointGroup, PrincipalType.DistributionList] : [PrincipalType.User]}
                  resolveDelay={250}
                  disabled={this.state.isSaving || this.state.isLoading} />
                <Label htmlFor={spField.InternalName}>{spField.Description}</Label>
              </div>
            );
          }

          if (spField.TypeAsString === 'UserMulti') {
            return (
              <div id={spField.InternalName}>
                <PeoplePicker
                  context={this.props.context}
                  personSelectionLimit={200}
                  titleText={spField.Title}
                  isRequired={spField.Required}
                  ensureUser={true}
                  selectedItems={(items) => { this.SavePeoplePickerMultiValue(`${spField.InternalName}`, items); }}
                  principalTypes={spField.SelectionMode === 1 ?
                    [PrincipalType.User, PrincipalType.SecurityGroup, PrincipalType.SharePointGroup, PrincipalType.DistributionList] : [PrincipalType.User]}
                  resolveDelay={250}
                  disabled={this.state.isSaving || this.state.isLoading} />
                <Label htmlFor={spField.InternalName}>{spField.Description}</Label>
              </div>
            );
          }

          if (spField.TypeAsString === 'TaxonomyFieldType') {
            return (
              <div id={spField.InternalName}>
                <Label required={spField.Required}>{spField.Title}</Label>
                <TaxonomyPicker
                  termsetNameOrID={spField.TermSetId}
                  panelTitle={spField.Title}
                  label={''}
                  context={this.props.context}
                  onChange={(value) => this.SaveTaxonomyFieldValue(spField.InternalName, value)}
                  isTermSetSelectable={false}
                  disabled={this.state.isSaving || this.state.isLoading}
                />
                <Label htmlFor={spField.InternalName}>{spField.Description}</Label>
              </div>
            );
          }

          if (spField.TypeAsString === 'TaxonomyFieldTypeMulti') {
            return (
              <div id={spField.InternalName}>
                <Label required={spField.Required}>{spField.Title}</Label>
                <TaxonomyPicker
                  allowMultipleSelections={true}
                  termsetNameOrID={spField.TermSetId}
                  panelTitle={spField.Title}
                  label={''}
                  context={this.props.context}
                  onChange={(value) => this.SaveTaxonomyMultiFieldValue(`${spField.InternalName}_0`, value)}
                  isTermSetSelectable={false}
                  disabled={this.state.isSaving || this.state.isLoading}
                />
                <Label htmlFor={spField.InternalName}>{spField.Description}</Label>
              </div>
            );
          }

          if (spField.TypeAsString === 'Lookup') {
            return (
              <div id={spField.InternalName}>
                <Label required={spField.Required}>{spField.Title}</Label>
                <ListItemPicker
                  listId={spField.LookupList}
                  columnInternalName={spField.LookupField}
                  itemLimit={1}
                  onSelectedItem={(value) => this.SaveLookupColumnValue(`${spField.InternalName}`, value)}
                  context={this.props.context}
                  disabled={this.state.isSaving || this.state.isLoading}
                />
                <Label htmlFor={spField.InternalName}>{spField.Description}</Label>
              </div>
            );
          }

          if (spField.TypeAsString === 'LookupMulti') {
            return (
              <div id={spField.InternalName}>
                <Label required={spField.Required}>{spField.Title}</Label>
                <ListItemPicker
                  listId={spField.LookupList}
                  columnInternalName={spField.LookupField}
                  itemLimit={200}
                  onSelectedItem={(value) => this.SaveLookupMultiColumnValue(`${spField.InternalName}`, value)}
                  context={this.props.context}
                  disabled={this.state.isSaving || this.state.isLoading} />
                <Label htmlFor={spField.InternalName}>{spField.Description}</Label>
              </div>
            );
          }
        }
      }).concat(
        <div className={styles.formButtonsContainer}>
          <PrimaryButton
            text={strings.SubmitButtonText}
            ariaDescription={strings.SubmitButtonText}
            onClick={() => this.SaveSiteRequest()}
            disabled={this.state.isSaving || this.state.isLoading || this.state.aliasValidating || !this.state.aliasIsValid || this.state.titleValidating || !this.state.titleIsValid}
          />
          <DefaultButton
            text={strings.CancelButtonText}
            ariaDescription={strings.CancelButtonText}
            onClick={() => this.ClearSiteRequest()}
            disabled={this.state.isSaving || this.state.isLoading || this.state.aliasValidating || this.state.titleValidating}
          />
        </div>
      )

      : <MessageBar messageBarType={MessageBarType.warning}>{strings.NoFieldsText}</MessageBar>;
  }


  // ***********
  // VALIDATION
  // ***********

  private ValidateRequiredField(value: any, required: boolean): string {
    if (required) {
      return value ? '' : strings.RequiredFieldMessage;
    }
    return '';
  }

  private ValidateNumericField(value: any, required: boolean, fieldType: string): string {
    if (required) {
      return this.ValidateRequiredField(value, required);
    }

    if (fieldType === 'Currency' && value) {
      let regex: RegExp = /^\d+(?:\.\d{0,2})$/;
      return !regex.test(value) ? strings.CurrencyFieldMessage : '';
    }

    return isNaN(value) ? strings.NumericFieldMessage : '';
  }

  private ValidateSubmit(postData: {}): boolean {

    if (this.props.siteProvisioningApiUrl && (!postData || !postData['ContentTypeId'] || !postData['EUMDivision'] || !postData['EUMSiteTemplate'])) {
      return false;
    }
    if (!this.props.siteProvisioningApiUrl && (!postData || !postData['ContentTypeId'] || !postData['EUMDivisionId'] || !postData['EUMSiteTemplateId'])) {
      return false;
    }

    let invalidFields: string[] = [];
    this.state.fields.forEach(spField => {
      if (!spField.Hidden) {
        if (spField.Required && !postData[spField.InternalName]) {
          invalidFields.push(spField.Title);
        }
      }
    });

    return invalidFields.length > 0 ? false : true;
  }

  public async CheckSiteExists(siteUrl: string): Promise<boolean> {
    try {
      // Make new web from url    
      const web = new Web(siteUrl);

      // Try to get web and only select Title
      const webWithTitle = await web.select('Title').get();

      // If web does exist make a return object and return
      if (webWithTitle.Title.length > 0) {
        return true;
      }

    }
    catch (error) {
      // if 404, it doesn't exist. Assume any other error means it exists
      const exists = error.status === 404 ? false : true;
      return exists;
    }
  }

  private async CheckIfGroupAliasValid(alias: string): Promise<boolean> {
    const spClient = new SPHttpClient();
    const resolve = spClient.get(`${this.props.tenantUrl}/_api/SP.Directory.DirectorySession/ValidateGroupName(displayName='${alias}',%20alias='${alias}')`)
      .then((response: Response) => {
        return response.json();
      }).then(result => {
        return result.d.ValidateGroupName.IsValidName;
      });
    return resolve;
  }

  private async ValidateAliasField(value: any, required: boolean): Promise<string> {
    if (value) {
      this.setState({ aliasValidating: true });
      let aliasValue= (`${this.state.selectedPrefix}${value}`).replace(/ /g, '-');
      value = `${aliasValue}`;

      // check the blacklist to make sure the entered title isn't blacklisted
      let blacklistedWord: string = this.CheckForBlacklistedWords(value);
      if (blacklistedWord) {
        // set the titleIsValid flag to false
        this.setState({ aliasValidating: false, aliasIsValid: false });
        return `${strings.BlacklistedWordsMessage} "${blacklistedWord}".`;
      }

      // check if the group alias is valid
      let groupAliasValid: boolean = await this.CheckIfGroupAliasValid(value);
      this.setState({ aliasIsValid: groupAliasValid });
      if (!this.state.aliasIsValid) {
        this.setState({ aliasValidating: false });
        return strings.AliasInUseMessage;
      }

      // check if the site URL is in use
      let siteUrl: string = `${this.props.tenantUrl}/sites/${value}`;
      let siteExists: boolean = await this.CheckSiteExists(siteUrl);
      this.setState({ aliasIsValid: !siteExists });
      if (!this.state.aliasIsValid) {
        this.setState({ aliasValidating: false });
        return strings.AliasInUseMessage;
      }

      this.setState({ aliasValidating: false });
    }

    if (required) {
      return value ? '' : strings.RequiredFieldMessage;
    }

    return '';
  }

  private async ValidateSiteUrlField(value: any, required: boolean): Promise<string> {
    if (value) {
      this.setState({ aliasValidating: true });
      let selectedPrefixValue: string = '';
      if (this.state.selectedPrefix) {
        selectedPrefixValue = this.state.selectedPrefix;
      }
      let aliasValue= (`${selectedPrefixValue}${value}`).replace(/ /g, '-');
      value = `/sites/${aliasValue}`;

      // check the blacklist to make sure the entered title isn't blacklisted
      // get everything after the last /
      let aliasPortion = value.indexOf('/') > -1 ? value.substring(value.lastIndexOf('/') + 1) : value;
      let blacklistedWord: string = this.CheckForBlacklistedWords(aliasPortion);
      if (blacklistedWord) {
        // set the titleIsValid flag to false
        this.setState({ aliasValidating: false, aliasIsValid: false });
        return `${strings.BlacklistedWordsMessage} "${blacklistedWord}".`;
      }

      // check if the site URL is in use
      let siteUrl: string = value;
      if (siteUrl.toLocaleLowerCase().indexOf(this.props.tenantUrl.toLocaleLowerCase()) === -1) {
        siteUrl = `${this.props.tenantUrl}${siteUrl}`;
      }

      let siteExists: boolean = await this.CheckSiteExists(siteUrl);
      this.setState({ aliasIsValid: !siteExists });
      if (!this.state.aliasIsValid) {
        this.setState({ aliasValidating: false });
        return strings.SiteUrlInUseMessage;
      }

      this.setState({ aliasValidating: false });
    }

    if (required) {
      return value ? '' : strings.RequiredFieldMessage;
    }

    return '';
  }

  private ValidateTitleField(value: any, required: boolean): string {
    if (value) {
      this.setState({ titleValidating: true });

      // check the blacklist to make sure the entered title isn't blacklisted
      let blacklistedWord: string = this.CheckForBlacklistedWords(value);
      if (blacklistedWord) {
        // set the titleIsValid flag to false
        this.setState({ titleValidating: false, titleIsValid: false });
        return `${strings.BlacklistedWordsMessage} "${blacklistedWord}".`;
      }

      // title doesn't contain any blacklisted characters
      this.setState({ titleValidating: false, titleIsValid: true });
    }

    if (required) {
      return value ? '' : strings.RequiredFieldMessage;
    }

    return '';
  }

  private CheckForBlacklistedWords(phrase: string): string {

    // if the blacklisted words list is null or empty, return null
    if (!this.state.blacklistedWords || this.state.blacklistedWords.length == 0 || !phrase) {
      return null;
    }

    // check the blacklisted words list and return the match if any found
    let blacklistedWord = null;
    this.state.blacklistedWords.some((value) => {
      // create a regex for this blacklisted word and perform a whole-word test (\b = word boundary, i = case insensitve)
      let wordsToSearch: RegExp = new RegExp('\\b' + value + '\\b', 'gi');

      // test the user-entered phrase to see if it contains any matches to this blacklisted word
      if (wordsToSearch.test(phrase)) {
        blacklistedWord = value;
        return true;
      }

      // additionally test the phrase with dashes replaced with spaces since we auto-generate the alias with spaces replaced with dashes
      if (wordsToSearch.test(phrase.replace(/-/g, ' '))) {
        blacklistedWord = value;
        return true;
      }

      // test the phrase with all non-alphanumeric characters removed to check if user tried to bypass blacklisted word
      if (wordsToSearch.test(phrase.replace(/[\W_]+/g, ' '))) {
        blacklistedWord = value;
        return true;
      }
    });

    return blacklistedWord;
  }

  // ***********
  // SAVING
  // ***********
  private async SaveTitleFieldValue(fieldInternalName: string, newValue: any, fieldType: string) {
    // remove leading/trailing spaces and invalid characters
    // update the alias
    if (newValue) {
      let generatedAlias: string = newValue.replace(/ /g, '-');

      // if not a modern team site, then site URL is being suggested instead so preppend /sites
      if (!this.ModernTeamSite) {
        generatedAlias = `${generatedAlias}`;
      }
      this.setState({ alias: generatedAlias, title: newValue });
      
    } else {
      this.setState({ alias: '' , title: ''});
    }
  }

  private SaveAliasFieldValue(newValue: any) {
    if (newValue) {
      let generatedAlias: string = newValue.replace(/ /g, '-');
      this.setState({ alias: generatedAlias });
    } else {
      this.setState({ alias: '' });
    }
  }

  private SaveFieldValue(fieldInternalName: string, newValue: any, fieldType: string) {
    let updateRequired: boolean = this.fieldValues.some(f => f.InternalName === fieldInternalName);

    if (!updateRequired) {
      let fieldValue: IFieldValue = { InternalName: fieldInternalName, Value: newValue };
      this.fieldValues.push(fieldValue);
    } else {
      this.fieldValues = this.fieldValues.map(f => {
        if (f.InternalName === fieldInternalName) {
          return { InternalName: fieldInternalName, Value: newValue };
        } else {
          return f;
        }
      });
    }
  }

  private SaveHtmlFieldValue = (fieldInternalName: string, newValue: any): string => {
    let updateRequired: boolean = this.fieldValues.some(f => f.InternalName === fieldInternalName);

    if (!updateRequired) {
      let fieldValue: IFieldValue = { InternalName: fieldInternalName, Value: newValue };
      this.fieldValues.push(fieldValue);
    } else {
      this.fieldValues = this.fieldValues.map(f => {
        if (f.InternalName === fieldInternalName) {
          return { InternalName: fieldInternalName, Value: newValue };
        } else {
          return f;
        }
      });
    }

    return newValue;
  }

  private SaveUrlFieldValue(fieldInternalName: string, value: any, fieldType: string) {
    let updateRequired: boolean = this.fieldValues.some(f => f.InternalName === fieldInternalName);

    let newValue: any;
    // the JSON is different depending if using API or SharePoint to submit
    if (value && this.props.siteProvisioningApiUrl) {
      newValue = {
        "value": value,
        "type": "Url"
      };
    } else if (value) {
      newValue = {
        "__metadata": { "type": "SP.FieldUrlValue" },
        "Description": value,
        'Url': value
      };
    }

    if (!updateRequired) {
      let fieldValue: IFieldValue = { InternalName: fieldInternalName, Value: newValue };
      this.fieldValues.push(fieldValue);
    } else {
      this.fieldValues = this.fieldValues.map(f => {
        if (f.InternalName === fieldInternalName) {
          return { InternalName: fieldInternalName, Value: newValue };
        } else {
          return f;
        }
      });
    }
  }

  private SaveDateFieldValue(fieldInternalName: string, newValue: Date) {
    let updateRequired: boolean = this.fieldValues.some(f => f.InternalName === fieldInternalName);

    let isoDateFormat: string = newValue ? new Date(newValue.getTime() + newValue.getTimezoneOffset() * 60000).toISOString() : "";

    if (!updateRequired) {
      let fieldValue: IFieldValue = { InternalName: fieldInternalName, Value: isoDateFormat };
      this.fieldValues.push(fieldValue);
    } else {
      this.fieldValues = this.fieldValues.map(f => {
        if (f.InternalName === fieldInternalName) {
          return { InternalName: fieldInternalName, Value: isoDateFormat };
        } else {
          return f;
        }
      });
    }
  }

  private SaveMultiSelectValue(fieldInternalName: string, newValue: any) {
    let selected: boolean = newValue.selected;
    let updateRequired: boolean = this.fieldValues.some(f => f.InternalName === fieldInternalName);

    let selections: any;
    if (!updateRequired && selected) {
      selections = { results: [newValue.key] };

      if (this.props.siteProvisioningApiUrl) {
        selections.type = 'ChoiceMulti';
      }

      let fieldValue: IFieldValue = { InternalName: fieldInternalName, Value: selections };
      this.fieldValues.push(fieldValue);
    } else {
      this.fieldValues = this.fieldValues.map(f => {
        if (f.InternalName === fieldInternalName) {
          let currentValue: string[] = f.Value.results;
          let updatedValue: string[];

          updatedValue = currentValue.filter((v) => v !== newValue.key);

          if (selected) {
            updatedValue.push(`${newValue.key}`);
          }

          selections = { results: updatedValue };
          if (this.props.siteProvisioningApiUrl) {
            selections.type = 'ChoiceMulti';
          }

          return { InternalName: fieldInternalName, Value: selections };
        } else {
          return f;
        }
      });
    }
  }

  private SavePeoplePickerValue(fieldInternalName: string, value: IPersonaProps[]) {
    let updateRequired: boolean = this.fieldValues.some(f => f.InternalName === fieldInternalName);

    let newValue: any;
    if (this.props.siteProvisioningApiUrl) {
      newValue = {
        "value": (value && value.length > 0) ? value[0].id : "",
        "type": "Person"
      };
    } else {
      newValue = (value && value.length > 0) ? value[0].id : "";
      fieldInternalName = `${fieldInternalName}Id`;
    }

    if (!updateRequired) {
      let fieldValue: IFieldValue = { InternalName: fieldInternalName, Value: newValue };
      this.fieldValues.push(fieldValue);
    } else {
      this.fieldValues = this.fieldValues.map(f => {
        if (f.InternalName === fieldInternalName) {
          return { InternalName: fieldInternalName, Value: newValue };
        } else {
          return f;
        }
      });
    }
  }

  private SavePeoplePickerMultiValue(fieldInternalName: string, value: IPersonaProps[]) {
    let updateRequired: boolean = this.fieldValues.some(f => f.InternalName === fieldInternalName);

    let newValue: string[] = [];
    value.forEach((v) => newValue.push(v.id));

    let personMultiJson: any;
    if (this.props.siteProvisioningApiUrl) {
      personMultiJson = {
        "value": newValue.join(','),
        "type": "PersonMulti"
      };
    } else {
      personMultiJson = { results: newValue };
      fieldInternalName = `${fieldInternalName}Id`;
    }

    if (!updateRequired) {
      let fieldValue: IFieldValue = { InternalName: fieldInternalName, Value: personMultiJson };
      this.fieldValues.push(fieldValue);
    } else {
      this.fieldValues = this.fieldValues.map(f => {
        if (f.InternalName === fieldInternalName) {
          return { InternalName: fieldInternalName, Value: personMultiJson };
        } else {
          return f;
        }
      });
    }
  }

  private SaveTaxonomyFieldValue(fieldInternalName: string, value: IPickerTerms) {
    let updateRequired: boolean = this.fieldValues.some(f => f.InternalName === fieldInternalName);

    let newValue: any;
    if (value && value.length > 0 && this.props.siteProvisioningApiUrl) {
      newValue = {
        "Label": value[0].name,
        "TermGuid": value[0].key,
        "type": "Taxonomy"
      };
    } else if (value && value.length > 0) {
      newValue = {
        "__metadata": { "type": "SP.Taxonomy.TaxonomyFieldValue" },
        "Label": value[0].name,
        'TermGuid': value[0].key,
        'WssId': '-1'
      };
    }

    if (!updateRequired) {
      let fieldValue: IFieldValue = { InternalName: fieldInternalName, Value: newValue };
      this.fieldValues.push(fieldValue);
    } else {
      this.fieldValues = this.fieldValues.map(f => {
        if (f.InternalName === fieldInternalName) {
          return { InternalName: fieldInternalName, Value: newValue };
        } else {
          return f;
        }
      });
    }
  }

  private SaveTaxonomyMultiFieldValue(taxonomyNoteFieldName: string, value: IPickerTerms) {
    let fieldInternalName = this.state.fields.filter(f => f.Title === taxonomyNoteFieldName)[0].InternalName;

    let updateRequired: boolean = this.fieldValues.some(f => f.InternalName === fieldInternalName);

    let newValue: any = '';
    value.forEach((v) => newValue = (newValue) ? (`${newValue}#-1;#${v.name}|${v.key};`) : (`-1;#${v.name}|${v.key};`));
    // remove trailing ;
    newValue = newValue.substring(0, newValue.length - 1);

    // format the value for the site provisioning API
    if (this.props.siteProvisioningApiUrl) {
      newValue = {
        "value": newValue,
        "type": "TaxonomyMulti"
      };
    }

    if (!updateRequired) {
      let fieldValue: IFieldValue = { InternalName: fieldInternalName, Value: newValue };
      this.fieldValues.push(fieldValue);
    } else {
      this.fieldValues = this.fieldValues.map(f => {
        if (f.InternalName === fieldInternalName) {
          return { InternalName: fieldInternalName, Value: newValue };
        } else {
          return f;
        }
      });
    }
  }

  private SaveLookupColumnValue(fieldInternalName: string, value: any[]) {
    let updateRequired: boolean = this.fieldValues.some(f => f.InternalName === fieldInternalName);

    let newValue: any;
    if (this.props.siteProvisioningApiUrl) {
      newValue = {
        "value": (value && value.length > 0) ? value[0].key : "",
        "type": "Lookup"
      };
    } else {
      newValue = (value && value.length > 0) ? value[0].key : "";
      fieldInternalName = `${fieldInternalName}Id`;
    }


    if (!updateRequired) {
      let fieldValue: IFieldValue = { InternalName: fieldInternalName, Value: newValue };
      this.fieldValues.push(fieldValue);
    } else {
      this.fieldValues = this.fieldValues.map(f => {
        if (f.InternalName === fieldInternalName) {
          return { InternalName: fieldInternalName, Value: newValue };
        } else {
          return f;
        }
      });
    }
  }

  private SaveLookupMultiColumnValue(fieldInternalName: string, value: any[]) {
    let updateRequired: boolean = this.fieldValues.some(f => f.InternalName === fieldInternalName);

    let newValue: string[] = [];
    value.forEach((v) => newValue.push(v.key));

    let lookupMultiJson: any;
    if (this.props.siteProvisioningApiUrl) {
      lookupMultiJson = {
        "value": newValue.join(','),
        "type": "LookupMulti"
      };
    } else {
      lookupMultiJson = { results: newValue };
      fieldInternalName = `${fieldInternalName}Id`;
    }

    if (!updateRequired) {
      let fieldValue: IFieldValue = { InternalName: fieldInternalName, Value: lookupMultiJson };
      this.fieldValues.push(fieldValue);
    } else {
      this.fieldValues = this.fieldValues.map(f => {
        if (f.InternalName === fieldInternalName) {
          return { InternalName: fieldInternalName, Value: lookupMultiJson };
        } else {
          return f;
        }
      });
    }
  }

  private SaveSiteRequest(): void {
    if (this.props.siteProvisioningApiUrl) {
      this.SaveSiteRequestWithApi();
    } else {
      this.SaveSiteRequestWithSharePoint();
    }
  }

  private SaveSiteRequestWithSharePoint(): void {
    this.setState({ isSaving: true });
    let postData = {};

    this.fieldValues.forEach((f) => postData[f.InternalName] = f.Value);
    postData['EUMDivisionId'] = this.state.selectedDivision;
    postData['EUMSiteTemplateId'] = this.state.selectedSiteTemplate;
    postData['ContentTypeId'] = this.state.contentTypeId;

    // set the default values for the toggles if they are hidden
    if (!this.ShowSiteVisibilityDropdown || postData['EUMSiteVisibility'] == null) {
      postData['EUMSiteVisibility'] = this.SiteVisibilityDefault;
    }
    if (this.ModernTeamSite && (!this.ShowCreateTeamToggle || postData['EUMCreateTeam'] == null)) {
      postData['EUMCreateTeam'] = this.CreateTeamDefault;
    }
    if (this.ModernTeamSite && (!this.ShowCreateOneNoteToggle || postData['EUMCreateOneNote'] == null)) {
      postData['EUMCreateOneNote'] = this.CreateOneNoteDefault;
    }
    if (this.ModernTeamSite && (!this.ShowCreatePlannerToggle || postData['EUMCreatePlanner'] == null)) {
      postData['EUMCreatePlanner'] = this.CreatePlannerDefault;
    }

    if (!this.ModernTeamSite) {
      postData['EUMCreateTeam'] = false;
      postData['EUMCreateOneNote'] = false;
      postData['EUMCreatePlanner'] = false;
    }

    // append the prefix, if any, to the Title field and Alias or SiteURL fields
    let selectedPrefixValue: string = '';
    if (this.state.selectedPrefix) {
      selectedPrefixValue = this.state.selectedPrefix;
    }
    let titleValue : string = this.state.title;
    titleValue = titleValue.trim();
    titleValue = titleValue.replace(/[\&\[\]/\\#,+()$~!@^%.=|'":;*?<>{}\-]/g, '');
    postData['Title'] = `${selectedPrefixValue}${titleValue}`;

    let aliasValue= (`${selectedPrefixValue}${this.state.alias}`).replace(/ /g, '-');
    if (this.ModernTeamSite) {
      postData['EUMAlias'] = `${aliasValue}`;
    } else {
      postData['EUMSiteURL'] = `/sites/${aliasValue}`;
    }

    let fieldsValid: boolean = this.ValidateSubmit(postData);
    if (!fieldsValid || !this.state.aliasIsValid || !this.state.titleIsValid) {
      this.setState({ fieldsValid: fieldsValid, isSaving: false });
    } else {
      sp.web.lists
        .getByTitle(this.props.sitesListName)
        .items
        .add(postData)
        .then((items: ItemAddResult) => {
          this.setState({ saveSuccess: true, isSaving: false });
        }).catch((e: Error): void => {
          this.setState({ hasError: true, errorMessage: e.message, isSaving: false, fieldsValid: fieldsValid });
        });
    }
  }

  private SaveSiteRequestWithApi(): void {
    this.setState({ isSaving: true });
    let postData = {};

    this.fieldValues.forEach((f) => postData[f.InternalName] = f.Value);
    postData['EUMDivision'] = {
      "value": this.state.selectedDivision,
      "type": "Lookup"
    };

    postData['EUMSiteTemplate'] = {
      "value": this.state.selectedSiteTemplate,
      "type": "Lookup"
    };

    postData['ContentTypeId'] = this.state.contentTypeId;

    // set the default values for the toggles if they are hidden
    if (!this.ShowSiteVisibilityDropdown || postData['EUMSiteVisibility'] == null) {
      postData['EUMSiteVisibility'] = this.SiteVisibilityDefault;
    }
    if (this.ModernTeamSite && (!this.ShowCreateTeamToggle || postData['EUMCreateTeam'] == null)) {
      postData['EUMCreateTeam'] = this.CreateTeamDefault;
    }
    if (this.ModernTeamSite && (!this.ShowCreateOneNoteToggle || postData['EUMCreateOneNote'] == null)) {
      postData['EUMCreateOneNote'] = this.CreateOneNoteDefault;
    }
    if (this.ModernTeamSite && (!this.ShowCreatePlannerToggle || postData['EUMCreatePlanner'] == null)) {
      postData['EUMCreatePlanner'] = this.CreatePlannerDefault;
    }

    if (!this.ModernTeamSite) {
      postData['EUMCreateTeam'] = false;
      postData['EUMCreateOneNote'] = false;
      postData['EUMCreatePlanner'] = false;
    }

    // append the prefix, if any, to the Title field and Alias or SiteURL fields
    let selectedPrefixValue: string = '';
    if (this.state.selectedPrefix) {
      selectedPrefixValue = this.state.selectedPrefix;
    }

    let titleValue : string = this.state.title;
    titleValue = titleValue.trim();
    titleValue = titleValue.replace(/[\&\[\]/\\#,+()$~!@^%.=|'":;*?<>{}\-]/g, '');
    postData['Title'] = `${selectedPrefixValue}${titleValue}`;

    let aliasValue= (`${selectedPrefixValue}${this.state.alias}`).replace(/ /g, '-');
    if (this.ModernTeamSite) {
      postData['EUMAlias'] = `${aliasValue}`;
    } else {
      postData['EUMSiteURL'] = `/sites/${aliasValue}`;
    }

    let fieldsValid: boolean = this.ValidateSubmit(postData);

    if (!fieldsValid || !this.state.aliasIsValid || !this.state.titleIsValid) {
      this.setState({ fieldsValid: fieldsValid, isSaving: false });
    } else {
      this.props.AadTokenProvider.getToken(this.props.siteProvisioningApiClientID).then((accessToken: string): void => {
        this.props.HttpClient.post(`${this.props.siteProvisioningApiUrl}/Sites`, HttpClient.configurations.v1,
          {
            headers: {
              'accept': 'application/json',
              'Content-type': 'application/json',
              'authorization': `Bearer ${accessToken}`
            },
            body: JSON.stringify(postData)
          })
          .then((response: HttpClientResponse): Promise<any> => {
            return new Promise<any>((resolve, reject) => {
              if (response.ok && response.status == 200) {
                response.json().then((responseJson) => {
                  resolve(responseJson);
                });
              } else {
                response.text().then((responseText) => {
                  reject(responseText);
                });
              }
            });
          })
          .then(() => {
            this.setState({ saveSuccess: true, isSaving: false });
          }).catch((e: Error): void => {
            this.setState({ hasError: true, errorMessage: e.message, isSaving: false, fieldsValid: fieldsValid });
          });
      }).catch((e: Error): void => {
        this.setState({ hasError: true, errorMessage: e.message, isSaving: false, fieldsValid: fieldsValid });
      });
    }
  }

  private ClearSiteRequest(): void {
    this.ModernTeamSite = false;
    this.ShowCreateTeamToggle = true;
    this.ShowSiteVisibilityDropdown = true;
    this.SiteVisibilityDefault = 'Public';
    this.CreateTeamDefault = true;
    this.CreateOneNoteDefault = true;
    this.ShowCreateOneNoteToggle = true;
    this.CreatePlannerDefault = true;
    this.ShowCreatePlannerToggle = true;

    this.fieldValues = [];
    this.setState({
      saveSuccess: false,
      hasError: false,
      fieldsValid: true,
      isLoading: true,
      isSaving: false,
      fieldsLoaded: false,
      divisionsLoaded: false,
      siteTemplatesLoaded: false,
      selectedDivision: '',
      selectedSiteTemplate: '',
      contentTypeId: '',
      divisions: [],
      siteTemplates: [],
      fields: [],
      errorMessage: '',
      alias: '',
      title: '',
      aliasValidating: false,
      aliasIsValid: false
    });
  }
}
