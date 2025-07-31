import { ISpUser } from "./ISpUser";

export interface INominationDetailsByLegal {
  id: number;
  isEmpAgreementSignedByCEO: boolean;
  isSavedOnLocalDrive: boolean;
  reviewer: ISpUser;
  reviewDate: string;
  nominationId: number;
  title:string;
}
