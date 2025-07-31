import { IUserDetails } from "../../models/IUserDetails";
export interface IUserRolesPermissions {
    IsActiveUser(email: string): Promise<boolean>;
}
