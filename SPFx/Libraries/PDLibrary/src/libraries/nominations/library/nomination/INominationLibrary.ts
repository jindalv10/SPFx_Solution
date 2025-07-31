
import { IAllNominationDetails } from "../../models/IAllNominationDetails";
import { ISpUser } from "../../models/ISpUser";
import { IUserDetails } from "../../models/IUserDetails";

export interface INominationLibrary {
    getNominationDetails(nominationId: number, nominee: ISpUser, currentUser: IUserDetails): Promise<IAllNominationDetails>;
    saveNominationDetails(allNominationDetails: IAllNominationDetails, currentUser: IUserDetails, assignPermission: ISpUser[],  groupName: string[], isAddToGroup: boolean): Promise<boolean>;
    deleteFile(nominationDetails: IAllNominationDetails, currentUser: IUserDetails, subFolderName: string, fileName: string) : Promise<boolean>;
}
