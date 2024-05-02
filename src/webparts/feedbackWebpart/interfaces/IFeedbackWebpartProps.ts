import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IFeedbackWebpartProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
  buttonText: string;
  PromptText: string;
  thankyouMessage: string;
  ColorOfSavebutton : string;
  ButtonSize: string;
  OpenFeedbackList: string;
  openPropertyPane: () => void;
 
}

export interface IFeedbackWebpartState {
  feedbackComment: string;
  isOpenPanel: boolean;
  saveFeedbackButtonText: string;
  promptText: string;
  thankyouMessage: string;
  colorOfSaveFeedbackButton: string;
  buttonSize: any;
  feedbackListName: string;
  feedbackMessage: string;
  feedbackCommentError: string;
  configurationGroupName: string;
  IsOwner: any;

}

export interface IFeedbackWebpartWebPartProps {
  description: string;
  buttonText: string;
  PromptText: string;
  thankyouMessage: string;
  ColorOfSavebutton : string;
  ButtonSize: string;
  OpenFeedbackList: string;
  openPropertyPane: () => void;
}