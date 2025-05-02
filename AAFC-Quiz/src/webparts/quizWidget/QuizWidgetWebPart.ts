import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { type IPropertyPaneConfiguration, PropertyPaneTextField} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'QuizWidgetWebPartStrings';
import QuizWidget from './components/QuizWidget';
import { IQuizWidgetProps } from './components/IQuizWidgetProps';
import { IUserInfo } from './models';
import SPHelper from './common/SPHelper';
import { SPFI, spfi ,SPFx } from "@pnp/sp";
import { CalloutTriggers } from '@pnp/spfx-property-controls/lib/PropertyFieldHeader';
import { PropertyFieldToggleWithCallout } from '@pnp/spfx-property-controls/lib/PropertyFieldToggleWithCallout';

export interface IQuizWidgetWebPartProps {
  pollQuestions: any[];
  MsgAfterSubmission: string;
  BtnSubmitQuizAnswersText: string;
  ResponseMsgToUserOnWrongAnswer:string;
  ResponseMsgToUser: string;
  pollBasedOnDate: boolean;
  NoPollMsg: string;
}

export default class QuizWidgetWebPart extends BaseClientSideWebPart<IQuizWidgetWebPartProps> {

  private helper: SPHelper | undefined = undefined;
  private userinfo: IUserInfo | undefined = undefined;
  private _sp:SPFI | undefined = undefined;
  private pageLanguage: string = "";

  protected async onInit(): Promise<void> {
      this._sp = spfi().using(SPFx(this.context));
      await super.onInit();
      this.helper = new SPHelper(this.context);
      this.userinfo = await this.helper.getCurrentUserInfo();
      this.pageLanguage = this.getPageLanguage();
      console.log(this._sp);
  }

  public render(): void {
    const element: React.ReactElement<IQuizWidgetProps> = React.createElement(
      QuizWidget,
      {
        context: this.context,
        quizQuestions: this.properties.pollQuestions,
        SuccessfullVoteSubmissionMsg: this.properties.MsgAfterSubmission,
        ResponseMsgToUser: this.properties.ResponseMsgToUser,
        BtnSubmitQuizAnswersText: this.properties.BtnSubmitQuizAnswersText,
        ResponseMsgToUserOnWrongAnswer: this.properties.ResponseMsgToUserOnWrongAnswer,
        pollBasedOnDate: this.properties.pollBasedOnDate,
        NoPollMsg: this.properties.NoPollMsg,
        currentUserInfo: this.userinfo,
        openPropertyPane: this.openPropertyPane,
        webServerRelativeUrl:this.context.pageContext.web.serverRelativeUrl,
        userEmail: this.context.pageContext.user.email,
        pageLanguage: this.pageLanguage,

      }
    );

    ReactDom.render(element, this.domElement);
  }
  

  private getPageLanguage(): string {
    const currentUrl = window.location.href;
    const urlParts = currentUrl.split("/");

    // Loop through the URL parts and find a two-letter language code
    for (let i = 0; i < urlParts.length; i++) {
        const part = urlParts[i];
        if (/^[a-z]{2}$/.test(part)) { // Match two-letter language code
            return part;
        }
    }

    return "en"; // Default to "en" if not found
  }

  protected get disableReactivePropertyChanges() {
    return false;
  }

  protected onDispose(): void {
      ReactDom.unmountComponentAtNode(this.domElement);
  }


  private openPropertyPane = (): void => {
      this.context.propertyPane.open();
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
                            PropertyFieldToggleWithCallout('pollBasedOnDate', {
                                calloutTrigger: CalloutTriggers.Hover,
                                key: 'pollBasedOnDateFieldId',
                                label: strings.PollDateLabel,
                                calloutContent: React.createElement('div', {}, strings.PollDateCalloutText),
                                onText: 'Yes',
                                offText: 'No',
                                checked: this.properties.pollBasedOnDate
                            }),
                            PropertyPaneTextField('ResponseMsgToUser', {
                                label: strings.ResponseMsgToUserLabel,
                                description: strings.ResponseMsgToUserDescription,
                                maxLength: 150,
                                multiline: true,
                                rows: 3,
                                resizable: false,
                                placeholder: strings.ResponseMsgToUserPlaceholder,
                                value: this.properties.ResponseMsgToUser
                            }),
                            PropertyPaneTextField('ResponseMsgToUserOnWrongAnswer', {
                                label: strings.ResponseMsgToUserWrongLabel,
                                description: strings.ResponseMsgToUserWrongDescription,
                                maxLength: 150,
                                multiline: true,
                                rows: 3,
                                resizable: false,
                                placeholder: strings.ResponseMsgToUserPlaceholder,
                                value: this.properties.ResponseMsgToUserOnWrongAnswer
                            }),
                            PropertyPaneTextField('BtnSubmitQuizAnswersText', {
                                label: strings.BtnSumbitVoteLabel,
                                description: strings.BtnSumbitVoteDescription,
                                maxLength: 50,
                                multiline: false,
                                resizable: false,
                                placeholder: strings.BtnSumbitVotePlaceholder,
                                value: this.properties.BtnSubmitQuizAnswersText
                            }),
                           
                            PropertyPaneTextField('NoPollMsg', {
                                label: strings.NoPollMsgLabel,
                                description: strings.NoPollMsgDescription,
                                maxLength: 150,
                                multiline: true,
                                rows: 3,
                                resizable: false,
                                placeholder: strings.NoPollMsgPlaceholder,
                                value: this.properties.NoPollMsg
                            })
                        ]
                    }
                ]
            }
        ]
    };
  }
}

