import * as React from 'react';
import styles from './PollWidget.module.scss';
import * as strings from 'PollWidgetWebPartStrings';
import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";
import OptionsContainer from './OptionsContainer/OptionsContainer';
import MessageContainer from './MessageContainer/MessageContainer';
import { IQuestionDetails, IResponseDetails, IPollAnalyticsInfo } from '../models';
import SPHelper from '../common/SPHelper';
import { MessageScope } from '../common/enumHelper';
import * as _ from 'lodash';
import * as moment from 'moment';
import { IPollWidgetProps } from './IPollWidgetProps';
import { IPollWidgetState } from './IPollWidgetState';
import { PrimaryButton, ProgressIndicator } from '@fluentui/react';
import QuickPollChart from './ChartContainer/QuickPollChart';
import {
  Accordion,
  AccordionItem,
  AccordionItemHeading,
  AccordionItemButton,
  AccordionItemPanel,
}  from "@pnp/spfx-controls-react/lib/AccessibleAccordion";

export default class PollWidget extends React.Component<IPollWidgetProps, IPollWidgetState> {
  private helper: SPHelper | undefined = undefined;
  private disQuestionId: string;
  private displayQuestion: IQuestionDetails | null;
  constructor(props: IPollWidgetProps) {
    super(props);
    this.state = {
      listExists: false,
      PollQuestions: [],
      UserResponse: [],
      displayQuestionId: "",
      displayQuestion: null,
      enableSubmit: true,
      enableChoices: true,
      showOptions: false,
      showProgress: false,
      showChart: false,
      showChartProgress: false,
      PollAnalytics: undefined,
      showMessage: false,
      isError: false,
      MsgContent: "",
      showSubmissionProgress: false,
      currentPollResponse: ""
    };
    this.helper = new SPHelper(this.props.context);
  }

  public componentDidMount = () => {
    this.checkAndCreateList();
  }

  public componentDidUpdate = (prevProps: IPollWidgetProps) => {
    if (prevProps.pollQuestions !== this.props.pollQuestions || prevProps.pollBasedOnDate !== this.props.pollBasedOnDate) {
      this.setState({
        UserResponse: [],
        displayQuestion: null,
        displayQuestionId: ''
      }, () => {
        this.getQuestions(this.props.pollQuestions);
      });
    }
    if (prevProps.chartType !== this.props.chartType) {
      let newPollAnalytics: IPollAnalyticsInfo = this.state.PollAnalytics;
      newPollAnalytics.ChartType = this.props.chartType;
      this.setState({
        PollAnalytics: newPollAnalytics
      }, this.bindResponseAnalytics);
    }
  }

  private async checkAndCreateList() {
    this.helper = new SPHelper(this.props.context);
    let listCreated = await this.helper.checkListExists();
    if (listCreated) {
      this.setState({ listExists: true }, () => {
        this.props.pollBasedOnType ? this.getListQuestions() : this.getQuestions();
      });
    }
  }

  private getQuestions = (questions?: any[]) => {
    let pquestions: IQuestionDetails[] = [];
    let tmpQuestions: any[] = (questions) ? questions : (this.props.pollQuestions) ? this.props.pollQuestions : [];
    if (tmpQuestions && tmpQuestions.length > 0) {
      tmpQuestions.map((question) => {
        pquestions.push({
          Id: question.uniqueId,
          DisplayName: question.QTitle,
          Choices: question.QOptions,
          UseDate: question.QUseDate,
          StartDate: new Date(question.QStartDate),
          EndDate: new Date(question.QEndDate),
          MultiChoice: question.QMultiChoice,
          SortIdx: question.sortIdx
        });
      });
    }
    this.disQuestionId = this.getDisplayQuestionID(pquestions);
    this.setState({ PollQuestions: pquestions, displayQuestionId: this.disQuestionId, displayQuestion: this.displayQuestion }, this.bindPolls);
  }

  private async getListQuestions(): Promise<void> {
    let pquestions: IQuestionDetails[] = [];
   
    const questionsResult = await this.helper.getQuestions();

    if (questionsResult && questionsResult.length > 0) {
        for (const question of questionsResult) {
            // Initialize a question object
            let questionObj: any = {
                Id: question.Id,
                DisplayName: this.props.pageLanguage !== "fr" ? question.DisplayName: question.DisplayNameFr,
                Choices: this.props.pageLanguage !== "fr" ? question.Choices: question.ChoicesFr,
                UseDate: question.UseDate,
                StartDate: new Date(question.StartDate),
                EndDate: new Date(question.EndDate),
                MultiChoice: question.MultiChoice,
                SortIdx: question.SortIdx,
                UserResponse: {},
                PollAnalytics: {}
            };

            // Fetch user responses for this specific question
            let usersResponse = await this.helper?.getPollResponse(question.Id);
            let filRes = _.filter(usersResponse, (o) => { return o.UserID == this.props.currentUserInfo.ID; });

            let currentPollResponse = "";
            if (filRes.length > 0) {
                currentPollResponse = filRes[0].Response ? filRes[0].Response : filRes[0].MultiResponse.join(',');
            }
            
            // Add the user's response to the question object
            questionObj.UserResponse = {
                PollResponse: currentPollResponse,
                FullResponses: filRes
            };

            // Now calculate the analytics for this question
            let tmpUserResponse = usersResponse;
            let tempData: any;
            let qChoices: string[] = questionObj.Choices.split(',') ?? [];
            let finalData: any = [];

            // Count responses based on whether it's a single or multiple choice question
            if (!question.MultiChoice) {
                tempData = _.countBy(tmpUserResponse, 'Response');
            } else {
                let data: any = [];
                tmpUserResponse.map((res: any) => {
                    if (res.MultiResponse && res.MultiResponse.length > 0) {
                        res.MultiResponse.map((finres: any) => {
                            data.push({
                                "UserID": res.UserID,
                                "Response": finres.trim()
                            });
                        });
                    }
                });
                tempData = _.countBy(data, 'Response');
            }

            // Map over choices to get the response count for each
            qChoices.map((label) => {
                if (tempData[label.trim()] == undefined) {
                    finalData.push(0);
                } else {
                    finalData.push(tempData[label.trim()]);
                }
            });

            // Add PollAnalytics to the question object
            questionObj.PollAnalytics = {
                Labels: qChoices,
                ChartType: this.props.chartType,
                Question: questionObj.DisplayName,
                PollResponse: finalData
            };

            // Push the question with analytics and user response into the array
            pquestions.push(questionObj);
        }
    }
    this.setState({
      PollQuestions: pquestions
    }, this.bindPolls);

  }
  private getDisplayQuestionID = (questions?: any[]) => {
    let filQuestions: any[] = [];
    if (questions && questions.length > 0) {
      if (this.props.pollBasedOnDate) {
        filQuestions = _.filter(questions, (o) => { return moment().startOf('date') >= moment(o.StartDate) && moment(o.EndDate) >= moment().startOf('date'); });
      } else {
        filQuestions = _.orderBy(questions, ['SortIdx'], ['asc']);
        this.displayQuestion = filQuestions[0];
        return filQuestions[0].Id;
      }
      if (filQuestions.length > 0) {
        filQuestions = _.orderBy(filQuestions, ['SortIdx'], ['asc']);
        this.displayQuestion = filQuestions[0];
        return filQuestions[0].Id;
      } else {
        this.displayQuestion = null;
      }
    }
    return '';
  }

  private bindPolls = () => {
    this.setState({
      showProgress: (this.state.PollQuestions.length > 0) ? true : false,
      enableSubmit: true,
      enableChoices: true,
      showOptions: false,
      showChart: false,
      showChartProgress: false,
      PollAnalytics: undefined,
      showMessage: false,
      isError: false,
      MsgContent: "",
      showSubmissionProgress: false
    }, this.getAllUsersResponse);
  }

  private _onChange = (ev: React.FormEvent<HTMLElement | HTMLInputElement>, option: any, isMultiSel: boolean, pollId: string): void => {
    let existingResponse: IResponseDetails[] = [];
    // Get the current PollQuestions array from the state
     let updatedPollQuestions = [...this.state.PollQuestions];
     // Find the specific question by pollId (match the question Id)
     let question = updatedPollQuestions.find((q) => q.Id === pollId);
     if (question) {
      // Prepare the updated user response object
      let userResponse: IResponseDetails = {
          PollQuestionId: question.Id, // The current question's Id
          PollQuestion: question.DisplayName, // The current question's DisplayName
          PollResponse: !isMultiSel ? option.key : '', // Single-choice poll response
          UserID: this.props.currentUserInfo.ID, // Current user's ID
          UserDisplayName: this.props.currentUserInfo.DisplayName, // User's Display Name
          UserLoginName: this.props.currentUserInfo.LoginName, // User's Login Name
          PollMultiResponse: isMultiSel ? option.key : [], // Multi-choice poll response (array)
          IsMulti: isMultiSel // Whether it's a multi-choice poll or not
      };

      if(updatedPollQuestions.length > 0) {
        // Filter the existing responses for this question to see if the user already responded
        const isExistingResponse = this.getUserResponse(question.UserResponse.FullResponses);
        if (isExistingResponse.length > 0) {
            let newExistingResponse: IResponseDetails = {
              PollQuestionId: question.Id, // The current question's Id
              PollQuestion: question.DisplayName, // The current question's DisplayName
              PollResponse: !isMultiSel ? option.key : '', // Single-choice poll response
              UserID: this.props.currentUserInfo.ID, // Current user's ID
              UserDisplayName: this.props.currentUserInfo.DisplayName, // User's Display Name
              UserLoginName: this.props.currentUserInfo.LoginName, // User's Login Name
              PollMultiResponse: isMultiSel ? option.key : [], // Multi-choice poll response (array)
              IsMulti: isMultiSel // Whether it's a multi-choice poll or not
            };
            existingResponse.push(newExistingResponse);
          } else {
              // If it's single-choice, update the PollResponse field
              existingResponse.push(userResponse);
          }
      } else{ 
          // If no existing response exists, add the new response to the question's user responses
          existingResponse.push(userResponse);
          question.UserResponse.FullResponses.push(userResponse);
      }
      // Update the PollQuestions array in the state
      this.setState({
        ...this.state,
        showChart: false,
        UserResponse: existingResponse,
        enableSubmit: false
      }); // Call bindPolls after state update
    }
  }

  private _getSelectedKey = (): string => {
    let selKey: string = "";
    if (this.state.UserResponse && this.state.UserResponse.length > 0) {
      var userResponses = this.state.UserResponse;
      var userRes = this.getUserResponse(userResponses);
      if (userRes.length > 0) {
        selKey = userRes[0].PollResponse;
      }
    }
    return selKey;
  }

  private _submitVote = async () => {
    // Initial state reset
    this.setState({
      enableSubmit: false,
      enableChoices: false,
      showSubmissionProgress: false,
      isError: false,
      MsgContent: '',
      showMessage: false
    });
  
    const curUserRes = this.getUserResponse(this.state.UserResponse);
    
    if (curUserRes.length <= 0) {
      // Validation error when no response is found
      this.setState({
        MsgContent: strings.SubmitValidationMessage,
        isError: true,
        showMessage: true,
        enableSubmit: true,
        enableChoices: true
      });
    } else {
      // Start the submission process
      this.setState({
        enableSubmit: false,
        enableChoices: false,
        showSubmissionProgress: true,
        isError: false,
        MsgContent: '',
        showMessage: false
      });
  
      try {
        // Await for the response submission to complete first
        await this.helper?.submitResponse(curUserRes[0]);
  
        // Now update the state after successful submission
        this.setState({
          showSubmissionProgress: false,
          showMessage: true,
          isError: false,
          MsgContent: (this.props.SuccessfullVoteSubmissionMsg && this.props.SuccessfullVoteSubmissionMsg.trim()) ?
            this.props.SuccessfullVoteSubmissionMsg.trim() : strings.SuccessfullVoteSubmission,
          showChartProgress: true
        });
              
        this.getAllUsersResponse()
  
      } catch (err) {
        console.log(err);
        // Handle failure in submission
        this.setState({
          enableSubmit: true,
          enableChoices: true,
          showSubmissionProgress: false,
          showMessage: true,
          isError: true,
          MsgContent: strings.FailedVoteSubmission
        });
      }
    }
  };
  

  private getAllUsersResponse = async (): Promise<void> => {
    let usersResponse = await this.helper.getPollResponse((this.state.displayQuestionId) ? this.state.displayQuestionId : this.disQuestionId);
    var filRes = _.filter(usersResponse, (o) => { return o.UserID == this.props.currentUserInfo.ID; });
    if (filRes.length > 0) {
      this.setState({
        showChartProgress: true,
        showChart: true,
        showOptions: false,
        showProgress: false,
        UserResponse: usersResponse,
        currentPollResponse: filRes[0].Response ? filRes[0].Response : filRes[0].MultiResponse.join(',')
      }, this.bindResponseAnalytics);
    } else {
      this.setState({
        showProgress: false,
        showOptions: true,
        showChartProgress: false,
        showChart: false
      });
    }
  }

  private bindResponseAnalytics = () => {
    const { PollQuestions, displayQuestionId } = this.state;
    let tmpUserResponse: any = this.state.UserResponse;
  
    if (tmpUserResponse && tmpUserResponse.length > 0) {
      // Find the question with the matching displayQuestionId
      const updatedPollQuestions = PollQuestions.map((pollQuestion: any) => {
        if (pollQuestion.Id === displayQuestionId) {
          var tempData: any;
          let qChoices: string[] = pollQuestion?.Choices?.split(',') ?? [];
          var finalData: any = [];
  
          if (!pollQuestion?.MultiChoice) {
            tempData = _.countBy(tmpUserResponse, 'Response');
          } else {
            var data: any = [];
            tmpUserResponse.forEach((res: any) => {
              if (res.MultiResponse && res.MultiResponse.length > 0) {
                res.MultiResponse.forEach((finres: any) => {
                  data.push({
                    "UserID": res.UserID,
                    "Response": finres.trim()
                  });
                });
              }
            });
            tempData = _.countBy(data, 'Response');
          }
  
          // For each choice, calculate the count of responses
          qChoices.forEach((label) => {
            if (tempData[label.trim()] === undefined) {
              finalData.push(0);
            } else {
              finalData.push(tempData[label.trim()]);
            }
          });
  
          const pollAnalytics: IPollAnalyticsInfo = {
            Labels: qChoices,
            ChartType: this.props.chartType,
            Question: pollQuestion?.DisplayName,
            PollResponse: finalData
          };
  
          // Return the updated question with PollAnalytics
          return {
            ...pollQuestion,
            PollAnalytics: pollAnalytics
          };
        }
        // Return the unmodified question if it doesn't match displayQuestionId
        return pollQuestion;
      });
  
      // Update the state with the updated PollQuestions
      this.setState({
        showProgress: false,
        showOptions: false,
        showChartProgress: false,
        showChart: true,
        PollQuestions: updatedPollQuestions // Updated PollQuestions state
      });
    }
  };
  

  private getUserResponse(UserResponses: IResponseDetails[]): IResponseDetails[] {
    let retUserResponse: IResponseDetails[];
    retUserResponse = UserResponses.filter((res) => { return res.UserID == this.props.currentUserInfo.ID; });
    return retUserResponse;
  }

  private handleAccordionChange = (expandedItems: string[]) => {
    // Assuming you're interested in the first expanded item
    if (expandedItems.length > 0) {
      this.setState({
        displayQuestionId: expandedItems[0], // Get the first expanded item ID
      }, this.getAllUsersResponse);
    } else {
      // If no items are expanded, reset the displayQuestionId
      this.setState({
        displayQuestionId: '',
      });
    }
  };

  public render(): React.ReactElement<IPollWidgetProps> {
    const { pollQuestions, BtnSubmitVoteText, ResponseMsgToUser, NoPollMsg } = this.props;

    const { enableSubmit, currentPollResponse, showProgress, showSubmissionProgress, showChartProgress, PollQuestions,
       showChart, listExists} = this.state;
    const showConfig: boolean = !this.props.pollBasedOnType && (!pollQuestions || pollQuestions.length <= 0 && (!PollQuestions || PollQuestions.length <= 0)) ? true : false;
    
    let userResponseCaption: string = (ResponseMsgToUser && ResponseMsgToUser.trim()) ? ResponseMsgToUser.trim() : strings.DefaultResponseMsgToUser;
    let submitButtonText: string = (BtnSubmitVoteText && BtnSubmitVoteText.trim()) ? BtnSubmitVoteText.trim() : strings.BtnSumbitVote;
    let nopollmsg: string = (NoPollMsg && NoPollMsg.trim()) ? NoPollMsg.trim() : strings.NoPollMsgDefault;
    
    return (
      <div className={styles.pollWidget}>
         <Accordion onChange={this.handleAccordionChange}>
        {!listExists ? (
          <ProgressIndicator label={strings.ListCreationText} description={strings.PlsWait} />
        ) : (
            <>
              {showConfig &&
              <Placeholder iconName='Edit'
                  iconText={strings.PlaceholderIconText}
                  description={strings.PlaceholderDescription}
                  buttonLabel={strings.PlaceholderButtonLabel}
                  onConfigure={this.props.openPropertyPane} />
              }
              {showProgress && !showChart &&
                <ProgressIndicator label={strings.QuestionLoadingText} description={strings.PlsWait} />
              }
              {PollQuestions.length <= 0 &&
                <MessageContainer MessageScope={MessageScope.Info} Message={nopollmsg} />
              }
                {PollQuestions && PollQuestions.length > 0 &&
                    PollQuestions.map((poll) => (
                      <AccordionItem uuid={poll.Id}>
                        <div className="ms-Grid" dir="ltr">
                          <div className="ms-Grid-row">
                            <div className="ms-Grid-col ms-lg12 ms-md12 ms-sm12">
                              <div className="ms-textAlignLeft ms-font-m-plus ms-fontWeight-semibold">
                              <AccordionItemHeading>
                                <AccordionItemButton>
                                    {poll.DisplayName}
                                </AccordionItemButton>
                              </AccordionItemHeading>
                              </div>
                            </div>
                          </div>
                        <AccordionItemPanel>
                          <div className="ms-Grid-row">
                            <div className="ms-Grid-col ms-lg6 ms-md6 ms-sm6">
                              <div className="ms-textAlignLeft ms-font-m-plus ms-fontWeight-semibold">
                                <OptionsContainer 
                                  disabled={false} 
                                  multiSelect={poll.MultiChoice}
                                  selectedKey={this._getSelectedKey} 
                                  options={poll.Choices} 
                                  label="Pick One" 
                                  PollId={poll.Id}
                                  onChange={(ev: any, option: any, isMultiSel: any) => this._onChange(ev, option, isMultiSel, poll.Id)} 
                                />
                              </div>
                            </div>
                          </div>
                          <div className="ms-Grid-row">
                            <div className="ms-Grid-col ms-lg12 ms-md12 ms-sm12">
                              <div className="ms-textAlignCenter ms-font-m-plus ms-fontWeight-semibold">
                                <PrimaryButton disabled={enableSubmit} text={submitButtonText}
                                  onClick={this._submitVote.bind(this)} />
                              </div>
                            </div>
                          </div>

                          {showSubmissionProgress && !showChartProgress &&
                            <ProgressIndicator label={strings.SubmissionLoadingText} description={strings.PlsWait} />
                          }
                          
                          {showChartProgress && !showChart &&
                            <ProgressIndicator label="Loading the Poll analytics" description="Getting all the responses..." />
                          }
                          {showChart &&
                            <>
                              <QuickPollChart PollAnalytics={poll.PollAnalytics} />
                              <MessageContainer MessageScope={MessageScope.Info} Message={`${userResponseCaption}: ${currentPollResponse}`} />
                            </>
                          }
                         
                        </AccordionItemPanel>
                        </div>
                      </AccordionItem>
                  ))}
            </>
          )
        }
        </Accordion>
      </div>
    );
  }
}
