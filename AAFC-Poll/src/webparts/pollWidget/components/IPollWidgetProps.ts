import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IUserInfo } from "../models";
import { ChartType } from "@pnp/spfx-controls-react/lib/ChartControl";

export interface IPollWidgetProps {
  context: WebPartContext
  pollQuestions: any[];
  SuccessfullVoteSubmissionMsg: string;
  ResponseMsgToUser: string;
  BtnSubmitVoteText: string;
  pollBasedOnDate: boolean;
  pollBasedOnType:boolean;
  currentUserInfo: IUserInfo;
  NoPollMsg: string;
  openPropertyPane: () => void;
  userEmail: string;
  webServerRelativeUrl:string;
  chartType: ChartType;
  pageLanguage: string;

}
