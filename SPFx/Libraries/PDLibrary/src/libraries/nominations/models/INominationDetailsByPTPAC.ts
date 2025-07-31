import { ISpUser } from "./ISpUser";

export interface INominationDetailsByPTPAC {
    id: number;
    reviewer: ISpUser;
    ptpacChair: ISpUser;
    recommendation: string;
    reviewDueDate: string;
    reviewDate: string;
    recommendationSentDate: string;
    reviewerAssignmentDate: string;
    nominationId?: number;
    ptpacChairComments: string;
    internalReviewDueDate:string;
  }
