import { ISpUser } from "./ISpUser";

export interface INominationDetailsByLA {
    id: number;
    title:string;
    assignee: ISpUser;
    isEmployeeAgreementSigned?: boolean;
    isEmployeeNumberUpdated?: boolean;
    reviewDate?: string;
    reviewNotes?: string;
    employeeNumberReversedDate?: string;
    isEmployeeNumberReversed?: boolean;
    withdrawCompletionDate?: string;
    nominationId?: number;
  }
