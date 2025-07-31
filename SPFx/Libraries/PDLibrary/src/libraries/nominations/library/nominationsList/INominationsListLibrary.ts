import { IDiscipline } from "../../models/IDiscipline";
import { IEmployee, IEmployeeUpdateProperties } from "../../models/IEmployee";
import { INominationListViewItem } from "../../models/INominationListViewItem";
import { INominationReviewer } from "../../models/INominationReviewer";
import { IProfessionalDesignationDetailed } from "../../models/IProfessionalDesignation";
import { IUserDetails } from "../../models/IUserDetails";
export interface INominationsListLibrary {
    getNominationList(currentUser: IUserDetails): Promise<INominationListViewItem[]>;
    getQCDisciplineUsers(pdDisciplineVal: string): Promise<INominationReviewer[]>;
    updateNomineeEmployeeDetails(employeeUpdateObject: IEmployeeUpdateProperties): Promise<IEmployee>;
    getProfessionalDesignationsByFinanceUserId(financeUserID: number, disciplines?: IDiscipline[]): Promise<IProfessionalDesignationDetailed[]>;
}
