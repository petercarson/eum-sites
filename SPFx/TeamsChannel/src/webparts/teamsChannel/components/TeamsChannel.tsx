import * as React from 'react';
import styles from './TeamsChannel.module.scss';
import { ITeamsChannelProps } from './ITeamsChannelProps';
import { Fabric } from 'office-ui-fabric-react/lib/Fabric';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { PrimaryButton } from 'office-ui-fabric-react/lib/components/Button';
import { Spinner } from 'office-ui-fabric-react/lib/Spinner';
import { Text } from 'office-ui-fabric-react/lib/Text';
import { Stack, MessageBar, MessageBarType, Toggle } from 'office-ui-fabric-react';
import * as strings from 'TeamsChannelWebPartStrings';
import { ITeamsChannelState } from './ITeamsChannelState';
import { sp, ItemAddResult } from "@pnp/sp";
import { getTheme } from 'office-ui-fabric-react/lib/Styling';
import { IChannelTemplateListItem } from './IChannelTemplateListItem';


export default class TeamsChannel extends React.Component<ITeamsChannelProps, ITeamsChannelState> {
  private ChannelName: string = null;
  private ChannelDescription: string = null;
  private IsPrivateChannel: boolean = false;
  private CreateOneNoteSection: boolean = false;
  private CreatePlanner: boolean = false;

  private PrivateChannelDefault: boolean = true;
  private ShowPrivateChannelToggle: boolean = true;
  private CreateOneNoteDefault: boolean = true;
  private ShowCreateOneNoteToggle: boolean = true;
  private CreatePlannerDefault: boolean = true;
  private ShowCreatePlannerToggle: boolean = true;


  constructor(props: ITeamsChannelProps) {
    super(props);

    this.state = {
      hasError: false,
      saveSuccess: false,
      isLoading: false,
      isSaving: false,
      fieldsValid: true,
      channelTemplatesLoaded: false,
      channelTemplates: null
    };
  }

  public render(): React.ReactElement<ITeamsChannelProps> {
    let getTemplates: boolean = (!this.state.hasError && !this.state.channelTemplatesLoaded);
    let renderChannelTemplatesDropdown: boolean = (this.state.channelTemplatesLoaded);
    let renderFields: boolean = (this.state.selectedChannelTemplate != null);
    return (
      <Fabric id="TeamsChannelRequest" className={styles.teamsChannel}>
        {(this.props.WebPartTitle) ? <span className={styles.webpartTitle}>{this.props.WebPartTitle.replace("{SiteCollectionTitle}", this.props.SiteName)}</span> : ""}

        {(this.state.saveSuccess) ? this.RenderSuccess() : null}

        {getTemplates ? this.GetChannelTemplates() : null}
        <Fabric hidden={this.state.saveSuccess || !this.state.channelTemplatesLoaded}>
          {renderChannelTemplatesDropdown ? this.RenderChannelTemplatesDropdown() : null}
          {renderFields ? this.RenderFields() : null}
        </Fabric>

        {(this.state.isLoading || this.state.isSaving) ? this.RenderLoadingSpinner() : null}
        {(this.state.hasError) ? this.RenderErrors() : null}
        {(!this.state.fieldsValid) ? this.RenderInvalidFieldsMessage() : null}
      </Fabric>
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

  private GetChannelTemplates(): void {
    sp.web.lists
      .getByTitle(this.props.ChannelTemplatesListName)
      .items
      .get()
      .then((items: IChannelTemplateListItem[]): void => {
        // if only 1 item returned, then default to that item
        if (items.length === 1) {
          this.setState({ isLoading: false, channelTemplates: items, channelTemplatesLoaded: true, selectedChannelTemplate: items[0].Id });
        } else {
          this.setState({ isLoading: false, channelTemplates: items, channelTemplatesLoaded: true });
        }
      }).catch((e: Error): void => {
        this.setState({ isLoading: false, hasError: true, errorMessage: e.message });
      });
  }

  private RenderChannelTemplatesDropdown() {
    if (this.state.channelTemplates.length <= 1) {
      return null;
    }

    return (
      <div id="ChannelTemplateDropdownSection">
        <Dropdown
          title="ChannelTemplate"
          id="ChannelTemplate"
          label={strings.ChannelTemplateDropdownLabel}
          placeholder={strings.ChannelTemplateDropdownPlaceholderText}
          required={true}
          options={this.state.channelTemplates.map(channelTemplate => ({ key: channelTemplate.Id, text: channelTemplate.Title }))}
          onChanged={(item) => this.setState({ selectedChannelTemplate: item.key.toString() })}
          disabled={this.state.isSaving || this.state.isLoading}
        />
      </div>
    );
  }

  private RenderFields() {
    this.CreateOneNoteDefault = this.state.channelTemplates.filter(s => s.Id == this.state.selectedChannelTemplate)[0].CreateOneNoteDefaultValue;
    this.ShowCreateOneNoteToggle = this.state.channelTemplates.filter(s => s.Id == this.state.selectedChannelTemplate)[0].CreateOneNoteShowToggle;
    this.CreateOneNoteSection = this.CreateOneNoteDefault;

    this.CreatePlannerDefault = this.state.channelTemplates.filter(s => s.Id == this.state.selectedChannelTemplate)[0].CreatePlannerDefaultValue;
    this.ShowCreatePlannerToggle = this.state.channelTemplates.filter(s => s.Id == this.state.selectedChannelTemplate)[0].CreatePlannerShowToggle;
    this.CreatePlanner = this.CreatePlannerDefault;

    this.PrivateChannelDefault = this.state.channelTemplates.filter(s => s.Id == this.state.selectedChannelTemplate)[0].PrivateChannelDefaultValue;
    this.ShowPrivateChannelToggle = this.state.channelTemplates.filter(s => s.Id == this.state.selectedChannelTemplate)[0].PrivateChannelShowToggle;
    this.IsPrivateChannel = this.PrivateChannelDefault;

    return (
      <Stack>
        <TextField
          label={strings.TeamsChannelTitleFieldLabel}
          onChange={this.SaveTeamsChannelName}
          required={true}
          validateOnLoad={false}
          validateOnFocusOut={true}
          onGetErrorMessage={(value) => this.ValidateRequiredField(value, true)}
          disabled={this.state.isSaving}
          placeholder={strings.TeamsChannelTitleFieldPlaceholder}
        />

        <TextField
          label={strings.TeamsChannelDescriptionFieldLabel}
          onChange={this.SaveTeamsChannelDescription}
          multiline rows={3}
          required={false}
          validateOnLoad={true}
          validateOnFocusOut={true}
          onGetErrorMessage={(value) => this.ValidateRequiredField(value, false)}
          disabled={this.state.isSaving}
          placeholder={strings.TeamsChannelDescriptionFieldPlaceholder}
        />

        {this.ShowPrivateChannelToggle ?
          <Stack>
            <Dropdown
              placeholder={strings.TeamsChannelPrivacyFieldPlaceholder}
              label={strings.TeamsChannelPrivacyFieldLabel}
              defaultSelectedKey={this.PrivateChannelDefault ? strings.TeamsChannelPrivacyOptionPrivate : strings.TeamsChannelPrivacyOptionPublic}
              options={[
                { key: strings.TeamsChannelPrivacyOptionPublic, text: strings.TeamsChannelPrivacyOptionPublic },
                { key: strings.TeamsChannelPrivacyOptionPrivate, text: strings.TeamsChannelPrivacyOptionPrivate }
              ]}
              required={true}
              onChange={this.SaveTeamsChannelPrivacy}
              disabled={this.state.isSaving}
            />
            <Text nowrap block variant="small" styles={{ root: { color: getTheme().palette.neutralSecondary } }}>
              {strings.TeamsChannelPrivacyFieldWarning}
            </Text>
          </Stack>
          : null}

        {this.ShowCreateOneNoteToggle ?
          <Toggle
            id="CreateOneNoteSection"
            defaultChecked={this.CreateOneNoteDefault}
            label={strings.CreateOneNoteSectionToggleLabel}
            onText={strings.ToggleOnText}
            offText={strings.ToggleOffText}
            onChanged={(value) => this.SaveCreateOneNoteSection(value)}
            disabled={this.state.isSaving}
          />
          : null}

        {this.ShowCreatePlannerToggle ?
          <Toggle
            id="CreateChannelPlanner"
            defaultChecked={this.CreatePlannerDefault}
            label={strings.CreatePlannerToggleLabel}
            onText={strings.ToggleOnText}
            offText={strings.ToggleOffText}
            onChanged={(value) => this.SaveCreatePlanner(value)}
            disabled={this.state.isSaving}
          />
          : null}

        <Fabric className={styles.formButtonsContainer}>
          <PrimaryButton text={strings.SubmitButtonText} disabled={this.state.isSaving || this.state.saveSuccess} onClick={this.SaveTeamsChannelRequest} />
        </Fabric>
      </Stack>
    );
  }

  private ValidateRequiredField(value: any, required: boolean): string {
    if (required) {
      return value ? '' : strings.RequiredFieldMessage;
    }
    return '';
  }

  private SaveTeamsChannelName = (event: React.FormEvent<HTMLInputElement>, newValue?: string): void => {
    this.ChannelName = newValue;
  }

  private SaveTeamsChannelDescription = (event: React.FormEvent<HTMLInputElement>, newValue?: string): void => {
    this.ChannelDescription = newValue;
  }

  private SaveTeamsChannelPrivacy = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
    this.IsPrivateChannel = (item && item.key === strings.TeamsChannelPrivacyOptionPrivate);
  }

  private SaveCreateOneNoteSection(newValue: any) {
    this.CreateOneNoteSection = newValue;
  }

  private SaveCreatePlanner(newValue: any) {
    this.CreatePlanner = newValue;
  }

  private SaveTeamsChannelRequest = (): void => {
    // validate required fields
    if (!this.ChannelName) {
      this.setState({ hasError: true, errorMessage: strings.FieldsInvalidErrorText });
    } else {
      this.setState({ isSaving: true, hasError: false, errorMessage: null });

      let postData = {};
      postData['Title'] = this.ChannelName;
      postData['Description'] = this.ChannelDescription;
      postData['IsPrivate'] = this.IsPrivateChannel;
      postData['TeamSiteURL'] = this.props.SiteUrl;
      postData['CreateOneNoteSection'] = this.CreateOneNoteSection;
      postData['CreateChannelPlanner'] = this.CreatePlanner;
      postData['ChannelTemplateId'] = this.state.selectedChannelTemplate;

      sp.web.lists
        .getByTitle(this.props.TeamsChannelsListName)
        .items
        .add(postData)
        .then(() => {
          this.setState({ saveSuccess: true, isSaving: false });
        }).catch((e: Error): void => {
          this.setState({ hasError: true, errorMessage: e.message, isSaving: false });
        });
    }
  }
}
