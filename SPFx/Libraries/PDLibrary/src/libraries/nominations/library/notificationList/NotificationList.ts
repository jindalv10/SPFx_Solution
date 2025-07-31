import { IAllNominationDetails } from "../../models/IAllNominationDetails";
import { INotificationDetails } from "../../models/INotificationDetails";
import { IUserDetails } from "../../models/IUserDetails";
import SPService from "../SPService";
import { Mapper } from "../startup/Mapper";
import { INotificationList } from "./INotificationList";

export default class NotificationList extends SPService implements INotificationList {

    private clientContext: any = null;
    constructor(context: any) {
        super(context);
        this.clientContext = context.pageContext;

    }



    public async getNotificationList(itemTitle: string, currentUser: IUserDetails, allNominationDataInfo: IAllNominationDetails, PDDisciplineVal ?: string, PDStatusVal ?: string): Promise<INotificationDetails[]> {
        //const user: ISiteUserInfo = await this.getCurrentSPUser();
        if (itemTitle && currentUser) {
            let camlQueryString = `<View><Query><Where>`;
            camlQueryString += PDDisciplineVal ? `<And>`: "";
            camlQueryString += PDStatusVal ? `<And>`: "";
            camlQueryString += PDDisciplineVal ? `<Eq><FieldRef Name='notificationPDDiscipline' /><Value Type='Choice'>${PDDisciplineVal}</Value></Eq>` : "";
            camlQueryString += PDStatusVal ? `<Eq><FieldRef Name='notificationPDStatus' /><Value Type='Choice'>${PDStatusVal}</Value></Eq>`: "";
            camlQueryString +=  PDStatusVal ? `</And>`: "";
            camlQueryString += `<Eq><FieldRef Name='Title' /><Value Type='Text'>${itemTitle}</Value></Eq>`;
            camlQueryString += PDDisciplineVal ? `</And>`: "";
            camlQueryString += `</Where><OrderBy><FieldRef Name='Id' Ascending='True'/></OrderBy></Query></View>`;
            const notificationData = await this.spGetByCamlQuery(this.Constants.SP_LIST_NAMES.NotificationList, camlQueryString);
            if (notificationData) {
                return notificationData.map((element) => {
                    return Mapper.mapNotificationDetailsList(element, allNominationDataInfo);
                });

            }
        }
    }

    public async nominationEmail(body: string, postURL: string): Promise<boolean> {
      if (body && postURL && this.Constants.SEND_EMAIL) {
        try {
            await this.callPowerAutomate(body, postURL);
            return true;
        } catch (error) {
            console.error("Exception:", error);
        }
      }
      return false;
    }

    public async nominationAttachmentPermission(body: string, postURL: string): Promise<boolean> {
      if (body && postURL && this.Constants.BREAK_ATTACHMENT_PERMISSION) {
          try {
              await this.callPowerAutomate(body, postURL);
              return true;
          } catch (error) {
              console.error("Exception:", error);
          }
      }
      return false;
  }
}
