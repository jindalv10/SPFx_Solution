import SPService from "../SPService";
import { IUserRolesPermissions } from "./IUserRolesPermissions";

export default class UserRolesPermissions extends SPService implements IUserRolesPermissions {

    constructor(context: any) {
        super(context);
    }
    public async IsActiveUser(email: string): Promise<boolean> {
        let employmentStatus = await this.isUserProfileActive(email);
        return employmentStatus ? Promise.resolve(true) : Promise.resolve(false);
    }
}