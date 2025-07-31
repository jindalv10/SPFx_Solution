import { ISpUser } from "./ISpUser";

export interface IIntakeNomination {
  id: number;
  title: string;
  nominee: ISpUser;
  epNominators?: ISpUser[];
  nomineeOffice?: string;
  nomineePractice?: string;
  nomineeDiscipline?: string;
  nomineeDesignation?: string;
  isProductPerson?: boolean;
  pdDiscipline?: string;
  pdStatus?: string;
  pdSubcategory?: string[];
  intakeNotes?: string;
  nominationStatus: string;
  rpCertification: boolean;
  submissionDate?: string;
  draftDate?: string;
  reSubmissionDate?: string;
  proficientLanguage?: string[];
  isStatusGrantedAfter2016: boolean;
  nominator?: ISpUser;
  financeUserID?: string;
  grantDate?: Date;
  trackCandidateNominated?:string;
  references?:ISpUser[];
  billingCode: string;
}




