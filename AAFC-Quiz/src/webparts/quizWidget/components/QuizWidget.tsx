import * as React from 'react';
import styles from './QuizWidget.module.scss';
import type { IQuizWidgetProps } from './IQuizWidgetProps';
import SPHelper from '../common/SPHelper';
import { IQuizDetails, IResponseDetails } from '../models';
import { IQuizWidgetState } from './IQuizWidgetState';
import * as _ from 'lodash';
import {
  Accordion,
  AccordionItem,
  AccordionItemHeading,
  AccordionItemButton,
  AccordionItemPanel,
}  from "@pnp/spfx-controls-react/lib/AccessibleAccordion";
import { PrimaryButton, ProgressIndicator } from '@fluentui/react';
import { Placeholder } from '@pnp/spfx-controls-react';
import MessageContainer from './MessageContainer/MessageContainer';
import { MessageScope } from '../common/enumHelper';
import OptionsContainer from './OptionsContainer/OptionsContainer';
import * as strings from 'QuizWidgetWebPartStrings';
import * as moment from 'moment';

export default class QuizWidget extends React.Component<IQuizWidgetProps, IQuizWidgetState> {
  private helper: SPHelper | undefined = undefined;

  constructor(props: IQuizWidgetProps) {
    super(props);
    this.state = {
      listExists: false,
      QuizQuestions: [],
      UserResponse: [],
      displayQuestionId: "",
      displayQuestion: null,
      enableSubmit: true,
      enableChoices: true,
      showOptions: false,
      showProgress: false,
      showMessage: false,
      showAnswer: false,
      isError: false,
      MsgContent: "",
      showSubmissionProgress: false,
      currentQuizResponse: { AnswerVal: "", isEqual: false }
    };
    this.helper = new SPHelper(this.props.context);
  }

  public componentDidMount = () => {
    this.checkAndCreateList();
  }

  public componentDidUpdate = (prevProps: IQuizWidgetProps) => {
    if (prevProps.quizQuestions !== this.props.quizQuestions || prevProps.pollBasedOnDate !== this.props.pollBasedOnDate) {
      this.setState({
        UserResponse: [],
        displayQuestion: null,
        displayQuestionId: ''
      }, () => {
        this.getListQuestions();
      });
    }
  }

  private async checkAndCreateList() {
    this.helper = new SPHelper(this.props.context);
    let listCreated = await this.helper.checkListExists();
    if (listCreated) {
      this.setState({ listExists: true }, () => {
         this.getListQuestions() 
      });
    }
  }

  private async getListQuestions(): Promise<void> {
    let qquestions: IQuizDetails[] = [];
   
    const questionsResult = await this.helper?.getQuestions();

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
                Answers: this.props.pageLanguage !== "fr" ? question.Answers: question.AnswersFr,
                Explanations: this.props.pageLanguage !== "fr" ? question.Explanations: question.ExplanationsFr,
            };


            // Push the question with analytics and user response into the array
            qquestions.push(questionObj);
        }
    }
    const displayQuestions = this.getDisplayQuestionID(qquestions);

    this.setState({
      QuizQuestions: displayQuestions
    });

  }

  private getDisplayQuestionID = (questions?: IQuizDetails[]) => {
    let filQuestions: any = [];
    if (questions.length > 0) {
      if (this.props.pollBasedOnDate) {
        filQuestions = _.filter(questions, (o) => { return moment().startOf('date') >= moment(o.StartDate) && moment(o.EndDate) >= moment().startOf('date'); });
      } else {
        filQuestions = _.orderBy(questions, ['SortIdx'], ['asc']);
        return filQuestions;
      }
      if (filQuestions.length > 0) {
        filQuestions = _.orderBy(filQuestions, ['SortIdx'], ['asc']);
        return filQuestions;
      } else {
        return questions;
      }
    }
    return '';
  }
  

  private _onChange = (ev: React.FormEvent<HTMLElement | HTMLInputElement>, option: any, isMultiSel: boolean, pollId: string): void => {
    let existingResponse: IResponseDetails[] = [];
    // Get the current PollQuestions array from the state
     let updatedPollQuestions = [...this.state.QuizQuestions];
     // Find the specific question by pollId (match the question Id)
     let question = updatedPollQuestions.find((q) => q.Id === pollId);
     if (question) {
      // Prepare the updated user response object
      let userResponse: IResponseDetails = {
          QuizQuestionId: question.Id, // The current question's Id
          QuizQuestion: question.DisplayName, // The current question's DisplayName
          QuizResponse: !isMultiSel ? option.key : '', // Single-choice poll response
          UserID: this.props.currentUserInfo?.ID, // Current user's ID
          UserDisplayName: this.props.currentUserInfo?.DisplayName, // User's Display Name
          UserLoginName: this.props.currentUserInfo?.LoginName, // User's Login Name
          QuizMultiResponse: isMultiSel ? option.key : [], // Multi-choice poll response (array)
          IsMulti: isMultiSel // Whether it's a multi-choice poll or not
      };

      if(updatedPollQuestions.length > 0) {
        existingResponse.push(userResponse);
      } else{ 
          // If no existing response exists, add the new response to the question's user responses
          existingResponse.push(userResponse);
      }
      // Update the PollQuestions array in the state
      this.setState({
        ...this.state,
        UserResponse: existingResponse
      }); // Call bindPolls after state update
    }
  }

  private _getSelectedKey = (): string => {
    let selKey: string = "";
    if (this.state.UserResponse && this.state.UserResponse.length > 0) {
      var userResponses = this.state.UserResponse;
      var userRes = this.getUserResponse(userResponses);
      if (userRes.length > 0) {
        selKey = userRes[0].QuizResponse || '';
      }
    }
    return selKey;
  }

  private _submitQuiz = async () => {
    // Initial state reset
    this.setState({
      enableSubmit: false,
      enableChoices: false,
      showSubmissionProgress: false,
      isError: false,
      MsgContent: '',
      showMessage: false
    });
  
    const curUserRes = this.getUserResponseAnswers(this.state.UserResponse);
    
    if (!curUserRes.hasOwnProperty('isEqual')) {
      // Validation error when no response is found
      this.setState({
        MsgContent: strings.SubmitValidationMessage,
        isError: true,
        showMessage: true,
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
        showMessage: false,
        showAnswer:true,
        currentQuizResponse: curUserRes
      });
  
      try {
        // Await for the response submission to complete first
        //await this.helper?.submitResponse(curUserRes[0]);
  
        // Now update the state after successful submission
        this.setState({
          showSubmissionProgress: false,
          showMessage: true,
          isError: false,
          MsgContent: (this.props.SuccessfullVoteSubmissionMsg && this.props.SuccessfullVoteSubmissionMsg.trim()) ?
            this.props.SuccessfullVoteSubmissionMsg.trim() : strings.SuccessfullVoteSubmission,
        });  
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

  private arraysMultiEqual(arr1: any, arr2: any) {
    try {
      if(arr1 && !Array.isArray(arr1))
      {
        arr1 = arr1.split(",")
      }
      // Check for null/undefined inputs
      if (!Array.isArray(arr1) || !Array.isArray(arr2)) {
          throw new Error("Both inputs must be arrays.");
      }

      // If arrays have different lengths, return false immediately
      if (arr1.length !== arr2.length) return { isEqual: false, matchingValues: [] };

      // Find matching values without duplicates (ES5 version)
      var matches = [];
      for (var i = 0; i < arr1.length; i++) {
          if (arr2.indexOf(arr1[i]) !== -1 && matches.indexOf(arr1[i]) === -1) {
              matches.push(arr1[i]);
          }
      }

      return { isEqual: matches.length === arr1.length, matchingValues: matches };

    } catch (error) {
        return { error: error.message };
    }
  }



  private getUserResponseAnswers(UserResponses: IResponseDetails[]): any {
    const { QuizQuestions } = this.state;

    // Filter UserResponses for the current user
    const retUserResponse = UserResponses.filter((res) => res.UserID === this.props.currentUserInfo?.ID);

    if (retUserResponse.length === 0) {
      return { isEqual: false, message: "No answers found" }; // Return false and message if no responses
    }

    // Filter QuizQuestions based on the QuizQuestionId from the user's response
    const retUserQuestions = QuizQuestions.filter((q) => q.Id === retUserResponse[0].QuizQuestionId);

    if (retUserQuestions.length === 0) {
        return { isEqual: false, message: "No matching questions found" }; // Return false and message if no matching questions
    }


    const userResponse = retUserResponse[0];
    const userQuestion = retUserQuestions[0];

    if (userResponse.IsMulti) {
        const responseValue = this.arraysMultiEqual(userQuestion.Answers, userResponse.QuizMultiResponse);
        if (responseValue.isEqual) {
            return { AnswerVal: responseValue.matchingValues.join(','), isEqual: responseValue.isEqual };
        }
        return { AnswerVal: retUserResponse[0].QuizMultiResponse.join(','), isEqual: false };
    } else {
        // If it's a single choice question, directly compare
        const responseValue = userQuestion.Answers === userResponse.QuizResponse;
        if (responseValue) {
            return { AnswerVal: userQuestion.Answers, isEqual: responseValue };
        }
        return { AnswerVal: retUserResponse[0].QuizResponse, isEqual: false };

    }
  }

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
        showAnswer: false,
        UserResponse: [],
        enableSubmit: true,
      });
    } else {
      // If no items are expanded, reset the displayQuestionId
      this.setState({
        displayQuestionId: '',
        showAnswer: false,
        UserResponse: [],
      });
    }
  };

  public render(): React.ReactElement<IQuizWidgetProps> {
    const {ResponseMsgToUserOnWrongAnswer, BtnSubmitQuizAnswersText, ResponseMsgToUser, NoPollMsg } = this.props;

    const {enableSubmit, showAnswer, currentQuizResponse, showProgress, showSubmissionProgress, QuizQuestions, listExists} = this.state;
    const showConfig: boolean =  !QuizQuestions || QuizQuestions.length <= 0 ? true : false;
    
    let userResponseCaption: string = (ResponseMsgToUser && ResponseMsgToUser.trim()) ? ResponseMsgToUser.trim() : strings.DefaultResponseMsgToUser;
    let submitButtonText: string = (BtnSubmitQuizAnswersText && BtnSubmitQuizAnswersText.trim()) ? BtnSubmitQuizAnswersText.trim() : strings.BtnSumbitVote;
 
    let nopollmsg: string = (NoPollMsg && NoPollMsg.trim()) ? NoPollMsg.trim() : strings.NoPollMsgDefault;
    
    return (
      <div className={styles.quizWidget}>
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
              {showProgress &&
                <ProgressIndicator label={strings.QuestionLoadingText} description={strings.PlsWait} />
              }
              {QuizQuestions.length <= 0 &&
                <MessageContainer MessageScope={MessageScope.Info} Message={nopollmsg} />
              }
                {QuizQuestions && QuizQuestions.length > 0 &&
                    QuizQuestions.map((quiz) => (
                      <AccordionItem uuid={quiz.Id}>
                        <div className="ms-Grid" dir="ltr">
                          <div className="ms-Grid-row">
                            <div className="ms-Grid-col ms-lg12 ms-md12 ms-sm12">
                              <div className="ms-textAlignLeft ms-font-m-plus ms-fontWeight-semibold">
                              <AccordionItemHeading>
                                <AccordionItemButton>
                                    {quiz.DisplayName}
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
                                  disabled={!enableSubmit} 
                                  multiSelect={quiz.MultiChoice}
                                  selectedKey={this._getSelectedKey} 
                                  options={quiz.Choices} 
                                  label="Pick One" 
                                  PollId={quiz.Id}
                                  onChange={(ev: any, option: any, isMultiSel: any) => this._onChange(ev, option, isMultiSel, quiz.Id)} 
                                />
                              </div>
                            </div>
                          </div>
                          <div className="ms-Grid-row">
                            <div className="ms-Grid-col ms-lg12 ms-md12 ms-sm12">
                              <div className="ms-textAlignCenter ms-font-m-plus ms-fontWeight-semibold">
                                <PrimaryButton disabled={!enableSubmit} text={submitButtonText}
                                  onClick={this._submitQuiz.bind(this)} />
                              </div>
                            </div>
                          </div>

                          {showSubmissionProgress &&
                            <ProgressIndicator label={strings.SubmissionLoadingText} description={strings.PlsWait} />
                          }
                          
                         
                          {
                            showAnswer && currentQuizResponse && !currentQuizResponse.isEqual ?
                              <MessageContainer 
                                MessageScope={MessageScope.Failure} 
                                Message={`${(ResponseMsgToUserOnWrongAnswer.replace("<UserResponse>", currentQuizResponse.AnswerVal)).replace("<CorrectAnswer>",quiz.Answers)}`}
                                LongMessage={`${quiz.Explanations}`} 
                              />
                              :
                              showAnswer && currentQuizResponse && currentQuizResponse.isEqual && currentQuizResponse.AnswerVal && currentQuizResponse.AnswerVal.length > 0 && (
                                <MessageContainer 
                                  MessageScope={MessageScope.Success} 
                                  Message={`${userResponseCaption.replace("<UserResponse>", currentQuizResponse.AnswerVal)}`} 
                                  LongMessage={`${quiz.Explanations}`}
                                />
                              )
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


