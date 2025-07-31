import { IAttachment } from "./IAttachment";
import { IIntakeNomination } from "./IIntakeNomination";
import { INominationDetailsByLA } from "./INominationDetailsByLA";
import { INominationDetailsByLegal } from "./INominationDetailsByLegal";
import { INominationDetailsByPTPAC } from "./INominationDetailsByPTPAC";
import { INominationDetailsByQC } from "./INominationDetailsByQC";
import { IReferences } from "./IReferences";

export interface IAllNominationDetails {
  intakeNomination: IIntakeNomination;
  nominationDetailsByLA?: INominationDetailsByLA;
  nominationDetailsByLegal?: INominationDetailsByLegal;
  nominationDetailsByQC?: INominationDetailsByQC;
  nominationDetailsByPTPAC?: INominationDetailsByPTPAC;
  nominationAttachments?: IAttachment[];
  nominationReferences?:IReferences[];
}
