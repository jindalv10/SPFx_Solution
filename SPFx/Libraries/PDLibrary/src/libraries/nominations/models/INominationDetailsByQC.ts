import { ISpUser } from "./ISpUser";

export interface INominationDetailsByQC {
  id: number;
  reviewer: ISpUser;
  additionalReviewer?: ISpUser;
  reviewNotes?: string;
  reviewDate?: string;
  sentToScDate?: string;
  draftDate?: string;
  withdrawnDate?: string;
  sentToPTPACDate?: string;
  notificationRecipient?: string[];
  nominationId: number;
  sentForMoreDetails?: string;
  qcStatus: string;
  reviewerAssignmentDate?: string;
  granted?: string;
  endDate?: string;
  anyoneElse: string;
  addPracticeDirector: boolean;
  nominationPasses: boolean;
  referencesPassed:string;
  qarPassed:string;
}
