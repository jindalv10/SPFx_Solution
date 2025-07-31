import { ISpUser } from "./ISpUser";

export interface INominationListViewItem {
  id: number;
  status: string;
  nominee: ISpUser;
  epNominators: ISpUser[];
  pdStatus: string;
  pdDiscipline: string;
  nominator: ISpUser;
  submitted: Date;
  sendSCforVoteDate?: Date;
  nominationPasses: boolean;
  referencesPassed:string;
  qarPassed:string;
  PTPACDueDate:Date;
  PTPACInternalDueDate:Date;
  PTPACReviewer:ISpUser;
  Subcategory:string[];
}




