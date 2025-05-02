import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IUserInfo } from "../models";

export interface IQuizWidgetProps {
  context: WebPartContext
  quizQuestions: any[];
  SuccessfullVoteSubmissionMsg: string;
  ResponseMsgToUser: string;
  BtnSubmitQuizAnswersText: string;
  ResponseMsgToUserOnWrongAnswer:  string;
  pollBasedOnDate: boolean;
  currentUserInfo: IUserInfo | undefined;
  NoPollMsg: string;
  openPropertyPane: () => void;
  userEmail: string;
  webServerRelativeUrl:string;
  pageLanguage: string;

}
