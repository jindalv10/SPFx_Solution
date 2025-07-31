import { IAllNominationDetails } from "../../models/IAllNominationDetails";
import { INotificationDetails } from "../../models/INotificationDetails";
import { IUserDetails } from "../../models/IUserDetails";
export interface INotificationList {
    getNotificationList(title: string, actor: IUserDetails, allNomationDataInfo: IAllNominationDetails, PDDisciplineVal ?: string, PDStatusVal ?: string): Promise<INotificationDetails[]>;
    nominationEmail(body: string, postURL: string): Promise<boolean>;
    nominationAttachmentPermission(body: string, postURL: string): Promise<boolean>;
  }
