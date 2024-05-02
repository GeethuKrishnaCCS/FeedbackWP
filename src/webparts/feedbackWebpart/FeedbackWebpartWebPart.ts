import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneButton,
  PropertyPaneSlider,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'FeedbackWebpartWebPartStrings';
import FeedbackWebpart from './components/FeedbackWebpart';
import { IFeedbackWebpartProps, IFeedbackWebpartWebPartProps } from './interfaces';



export default class FeedbackWebpartWebPart extends BaseClientSideWebPart<IFeedbackWebpartWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<IFeedbackWebpartProps> = React.createElement(
      FeedbackWebpart,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        context: this.context,
        buttonText: this.properties.buttonText,
        PromptText: this.properties.PromptText,
        thankyouMessage: this.properties.thankyouMessage,
        ColorOfSavebutton: this.properties.ColorOfSavebutton,
        ButtonSize: this.properties.ButtonSize,
        OpenFeedbackList: this.properties.OpenFeedbackList,
        openPropertyPane: this.openPropertyPane.bind(this),

      }
    );

    ReactDom.render(element, this.domElement);
  }




  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }


  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              throw new Error('Unknown host');
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            },
            {
              groupName: "General",
              groupFields: [
                PropertyPaneButton('OpenFeedbackList', {
                  text: 'Open Feedback List',
                  onClick: () => { }

                }),
              ]
            },
            {
              groupName: "Change Labels & Language",
              groupFields: [
                PropertyPaneTextField('buttonText', {
                  label: 'Save Feedback Button Text'
                }),
                PropertyPaneTextField('PromptText', {
                  label: 'Prompt Text'
                }),
                PropertyPaneTextField('thankyouMessage', {
                  label: 'Thankyou Message'
                }),
              ]
            },
            {
              groupName: "Change Colors & Style",
              groupFields: [
                PropertyPaneTextField('ColorOfSavebutton', {
                  label: 'Color of "Save Feedback" Button'
                }),
                PropertyPaneSlider('ButtonSize', {
                  label: 'Button Size',
                  min: 100,
                  max: 500,
                  value: 100,
                  showValue: true,
                  step: 1
                }),

              ]
            },
          ]
        }
      ]
    };
  }


  public openPropertyPane() {
    this.context.propertyPane.openDetails();
   
  }

}


