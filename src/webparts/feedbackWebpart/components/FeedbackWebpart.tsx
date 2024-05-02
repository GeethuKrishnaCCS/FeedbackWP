import * as React from 'react';
import styles from './FeedbackWebpart.module.scss';
// import { escape } from '@microsoft/sp-lodash-subset';
import { IFeedbackWebpartProps, IFeedbackWebpartState } from '../interfaces';
import { DefaultButton, IIconProps, IconButton, MessageBar, MessageBarType, Panel, PrimaryButton, Slider, TextField } from '@fluentui/react';
import { FeedbackWebpartService } from '../services';
import Toast from './Toast';

export default class FeedbackWebpart extends React.Component<IFeedbackWebpartProps, IFeedbackWebpartState, {}> {
  private _service: any;

  public constructor(props: IFeedbackWebpartProps) {
    super(props);
    this._service = new FeedbackWebpartService(this.props.context);

    this.state = {
      feedbackComment: "",
      isOpenPanel: false,
      saveFeedbackButtonText: "",
      promptText: "",
      thankyouMessage: "",
      colorOfSaveFeedbackButton: "",
      feedbackListName: "",
      buttonSize: 100,
      feedbackMessage: "",
      feedbackCommentError: "",
      configurationGroupName: "",
      IsOwner: "",
    }

    this.onChangeComment = this.onChangeComment.bind(this);
    this.onSubmitClick = this.onSubmitClick.bind(this);
    this.onOpenPropertyPane = this.onOpenPropertyPane.bind(this);
    this.openPanel = this.openPanel.bind(this);
    this.hidePanel = this.hidePanel.bind(this);
    this.onSaveFeedbackButtonText = this.onSaveFeedbackButtonText.bind(this);
    this.onChangePromptText = this.onChangePromptText.bind(this);
    this.onChangeThankyouMessage = this.onChangeThankyouMessage.bind(this);
    this.onChangeColorOfSaveFeedbackButton = this.onChangeColorOfSaveFeedbackButton.bind(this);
    this.onSliderChange = this.onSliderChange.bind(this);
    this.onChangefeedbackListName = this.onChangefeedbackListName.bind(this);
    this.openFeedbackList = this.openFeedbackList.bind(this);
    this.enterYourFeedbackMessage = this.enterYourFeedbackMessage.bind(this);
    this.onClickOk = this.onClickOk.bind(this);
    // this.getFeedbackSettingsList = this.getFeedbackSettingsList.bind(this);
    this.getCurrentUser = this.getCurrentUser.bind(this);
    this.getSiteUser = this.getSiteUser.bind(this);
    this.getSiteGroup = this.getSiteGroup.bind(this);

    this.onConfigurationGroupName = this.onConfigurationGroupName.bind(this);
    // this.checkUserGroupMembership = this.checkUserGroupMembership.bind(this);
  }


  public async componentDidMount() {
    await this.getFeedbackSettingsList();
    await this.getCurrentUser();
    // await this.getSiteUser();
    // await this.getSiteGroup();
    await this.checkUserGroupMembership();
  }

  public async getCurrentUser() {
    const getcurrentuser = await this._service.getCurrentUser();
    console.log('getcurrentuser: ', getcurrentuser);
  }

  public async getSiteUser() {
    const getsiteuser = await this._service.getSiteUsers();
    console.log('getsiteuser: ', getsiteuser);
  }

  public async getSiteGroup() {
    const getSiteGroup = await this._service.getSiteGroups();
    console.log('getSiteGroup: ', getSiteGroup);
  }

  public async checkUserGroupMembership() {
    const isOwner = await this._service.isUserOwnerOfGroup();
    console.log('isOwner: ', isOwner);
    const userGroups = isOwner.some((group: any) => group.Title === this.state.configurationGroupName);
    console.log('this.state.configurationGroupName: ', this.state.configurationGroupName);
    // console.log('userGroups: ', userGroups);
    // "MaterialRequest Owners"
    this.setState({ IsOwner: userGroups });
    console.log("Is user a owner?", this.state.IsOwner);
  }




  public onChangeComment(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, feedbackcomment: string) {
    this.setState({ feedbackComment: feedbackcomment });
  }
  public onChangefeedbackListName(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, feedbackListName: string) {
    this.setState({ feedbackListName: feedbackListName });
    console.log('feedbackListName: ', feedbackListName);
    console.log('feedbackListNameState: ', this.state.feedbackListName);
  }
  public enterYourFeedbackMessage(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, feedbackMessage: string) {
    this.setState({ feedbackMessage: feedbackMessage });
  }
  public onSaveFeedbackButtonText(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, SaveFeedbackButtonText: string) {
    this.setState({ saveFeedbackButtonText: SaveFeedbackButtonText });
  }
  public onChangePromptText(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, PromptText: string) {
    this.setState({ promptText: PromptText });
  }
  public onChangeThankyouMessage(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, ThankyouMessage: string) {
    this.setState({ thankyouMessage: ThankyouMessage });
  }
  public onChangeColorOfSaveFeedbackButton(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, ColorOfSaveFeedbackButton: string) {
    this.setState({ colorOfSaveFeedbackButton: ColorOfSaveFeedbackButton });
  }

  public onConfigurationGroupName(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, ConfigurationGroupName: string) {
    this.setState({ configurationGroupName: ConfigurationGroupName });
  }

  public async getFeedbackSettingsList() {
    const url: string = this.props.context.pageContext.web.serverRelativeUrl;
    const items = await this._service.getListItems("FeedbackSettingsList", url);

    if (items && items.length > 0) {
      const feedbackSettings = items[0];
      this.setState({
        feedbackListName: feedbackSettings.FeedbackListName || "",
        feedbackMessage: feedbackSettings.EnterYourFeedbackMessage || "",
        promptText: feedbackSettings.PromptText || "",
        saveFeedbackButtonText: feedbackSettings.SaveFeedbackButtonText || "",
        thankyouMessage: feedbackSettings.ThankyouMessage || "",
        colorOfSaveFeedbackButton: feedbackSettings.ColorOfSaveFeedbackButton || "",
        buttonSize: feedbackSettings.ButtonSize || 150,
        configurationGroupName: feedbackSettings.ConfigurationGroupName || "",
      });
    }
  }

  public async onSubmitClick(): Promise<void> {
    if (!this.state.feedbackComment) {
      this.setState({ feedbackCommentError: this.state.promptText });
      // this.setState({ feedbackCommentError: 'Feedback is required.' });
      return;
    }

    const dataItem = {
      Feedback: this.state.feedbackComment,
    };

    const url: string = this.props.context.pageContext.web.serverRelativeUrl;
    this._service.addListItem(dataItem, this.state.feedbackListName, url).then(async (item: any) => {
      console.log('item: ', item);
    })
    Toast("success", this.state.thankyouMessage);
    this.setState({feedbackComment : ""})
  }

  public async onClickOk(): Promise<void> {
    const ItemList = {
      FeedbackListName: this.state.feedbackListName,
      PromptText: this.state.promptText,
      SaveFeedbackButtonText: this.state.saveFeedbackButtonText,
      ThankyouMessage: this.state.thankyouMessage,
      ColorOfSaveFeedbackButton: this.state.colorOfSaveFeedbackButton,
      ButtonSize: this.state.buttonSize,
      EnterYourFeedbackMessage: this.state.feedbackMessage,
      ConfigurationGroupName: this.state.configurationGroupName
    };

    const url: string = this.props.context.pageContext.web.serverRelativeUrl;
    const items = await this._service.getListItems("FeedbackSettingsList", url);

    if (items && items.length > 0) {
      const existingItem = items[0];
      const dataItem = {
        Id: existingItem.Id,
        ...ItemList,
      };
      await this._service.updateItem("FeedbackSettingsList", dataItem, existingItem.Id, url);
    } else {
      await this._service.addListItem(ItemList, "FeedbackSettingsList", url);
    }
    this.hidePanel();
  }


  public openFeedbackList() {
    const url: string = this.props.context.pageContext.web.serverRelativeUrl;
    // const feedbackListUrl = url + "/Lists/" + "/FeedbackList"
    const feedbackListUrl = url + "/Lists/" + this.state.feedbackListName;
    console.log('Redirecting to:', feedbackListUrl);

    // window.location.href = feedbackListUrl;
    window.open(feedbackListUrl, '_blank');
  }


  public onOpenPropertyPane(): void {
    this.props.context.propertyPane.openDetails();
  }

  public async openPanel(): Promise<void> {
    this.setState({ isOpenPanel: true });
  }

  public hidePanel = () => {
    this.setState({ isOpenPanel: false });
  };

  public onSliderChange = (value: number) => {
    this.setState({ buttonSize: value });
  };


  public render(): React.ReactElement<IFeedbackWebpartProps> {
    const OpenPropertypaneIcon: IIconProps = { iconName: 'Refresh' };

    const {

      hasTeamsContext,

    } = this.props;

    // const onRenderFooterContent = () => {
    //   return (
    //     <div>
    //       <PrimaryButton onClick={this.hidePanel}>
    //         Ok
    //       </PrimaryButton>
    //       <DefaultButton onClick={this.hidePanel}>Cancel</DefaultButton>
    //     </div>
    //   );
    // };

    return (
      <section className={`${styles.feedbackWebpart} ${hasTeamsContext ? styles.teams : ''}`}>
        {/* <div>
          <IconButton
            iconProps={OpenPropertypaneIcon}
            ariaLabel="Close popup modal"
            // className={styles.OpenPropertypaneIconbtn}
            onClick={this.onOpenPropertyPane}
          />
        </div> */}

        <div>
          {/* {this.state.IsOwner === true &&
            <div className={styles.OpenPropertypaneIconbtndiv}>
              <IconButton
                iconProps={OpenPropertypaneIcon}
                ariaLabel="Close popup modal"
                // className={styles.OpenPropertypaneIconbtn}
                onClick={this.openPanel}
              />
            </div>
          } */}

          {(!this.state.feedbackListName || !this.state.configurationGroupName) && (
            <div className={styles.OpenPropertypaneIconbtndiv}>
              <IconButton
                iconProps={OpenPropertypaneIcon}
                ariaLabel="Close popup modal"
                onClick={this.openPanel}
              />
            </div>
          )}

          {this.state.feedbackListName && this.state.configurationGroupName && this.state.IsOwner === true && (
            <div className={styles.OpenPropertypaneIconbtndiv}>
              <IconButton
                iconProps={OpenPropertypaneIcon}
                ariaLabel="Close popup modal"
                onClick={this.openPanel}
              />
            </div>
          )}



          <Panel
            headerText="Feedback to List"
            isOpen={this.state.isOpenPanel}
            onDismiss={this.hidePanel}
            closeButtonAriaLabel="Close"
          // onRenderFooterContent={onRenderFooterContent}
          >
            <DefaultButton
              text='Open Feedback List'
              onClick={this.openFeedbackList}
            />

            <TextField
              label="Feedback List Name"
              onChange={this.onChangefeedbackListName}
              placeholder={this.state.feedbackListName}
              value={this.state.feedbackListName}
            />

            <TextField
              label="Enter Your Feedback Message"
              onChange={this.enterYourFeedbackMessage}
              placeholder={this.state.feedbackMessage}
              value={this.state.feedbackMessage}
            />

            <TextField
              label="Prompt Text"
              onChange={this.onChangePromptText}
              placeholder={this.state.promptText}
              value={this.state.promptText}
            />
            <TextField
              label="Save Feedback Button Text"
              onChange={this.onSaveFeedbackButtonText}
              placeholder={this.state.saveFeedbackButtonText}
              value={this.state.saveFeedbackButtonText}
            />
            <TextField
              label="Thankyou Message"
              onChange={this.onChangeThankyouMessage}
              placeholder={this.state.thankyouMessage}
              value={this.state.thankyouMessage}
            />
            <TextField
              label='Color of "Save Feedback" Button'
              onChange={this.onChangeColorOfSaveFeedbackButton}
              placeholder={this.state.colorOfSaveFeedbackButton}
              value={this.state.colorOfSaveFeedbackButton}
            />

            <TextField
              label='Configuration Group Name'
              onChange={this.onConfigurationGroupName}
              placeholder={this.state.configurationGroupName}
              value={this.state.configurationGroupName}
            />

            <Slider
              label='Button Size'
              min={150}
              max={500}
              step={2}
              defaultValue={150}
              onChange={this.onSliderChange}
              value={this.state.buttonSize}
            />
            <div className={styles.btndivPanel}>
              <PrimaryButton
                text='OK'
                onClick={this.onClickOk}>

              </PrimaryButton>
              <DefaultButton onClick={this.hidePanel}>Cancel</DefaultButton>
            </div>

          </Panel>
        </div>

        <div>
          <TextField
            // label="Comment"
            // placeholder={this.props.PromptText}
            placeholder={this.state.feedbackMessage}
            multiline rows={3}
            onChange={this.onChangeComment}
            value={this.state.feedbackComment}
          // className={styles.commentArea}
          />

          {(!this.state.feedbackComment && this.state.feedbackCommentError) && (
            <div className={styles.msgbar}>
              <MessageBar
                messageBarType={MessageBarType.error}
                isMultiline={false}
                dismissButtonAriaLabel="Close"
              >
                {this.state.feedbackCommentError}
              </MessageBar>
            </div>
          )}

          <div className={styles.btndiv}>
            <PrimaryButton
              text={this.state.saveFeedbackButtonText}
              onClick={this.onSubmitClick}
              styles={{
                root: {
                  background: this.state.colorOfSaveFeedbackButton,
                  borderColor: this.state.colorOfSaveFeedbackButton,
                  width: this.state.buttonSize + 'px',
                }
              }}
            />
          </div>

        </div>

      </section>
    );
  }
}
