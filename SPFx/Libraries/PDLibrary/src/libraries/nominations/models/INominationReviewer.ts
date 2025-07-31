import { ISpUser } from "./ISpUser";

export interface INominationReviewer {
  PDDiscipline?: string;
  AuthorizedQC?: ISpUser;
  AuthorizedPTPAC?: ISpUser;
}
