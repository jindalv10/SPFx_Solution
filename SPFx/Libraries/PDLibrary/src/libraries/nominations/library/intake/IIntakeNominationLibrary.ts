import { IMasterDetails } from "../../models/IMasterDetails";
import { INomineeDetails } from "../../models/INomineeDetails";
import { INomineeExist } from "../../models/IUserDetails";


export interface IIntakeNominationLibrary {
    getNomineeDetailsFromEmpDB(nomineeEmailAddress: string): Promise<INomineeDetails>;
    getMasterDetails(): Promise<IMasterDetails>;
    checkIfValidNominee(nomineeId: number): Promise<INomineeExist>;
    checkIfValidNomineeWithDiscAndPDStatus(financeId: string, selectedNomineePDStatus: string, selectedNomineePDDiscipline: string): Promise<INomineeExist>;
    getEmployeeInformation(financeUserId:string): Promise<string[]>;
}

