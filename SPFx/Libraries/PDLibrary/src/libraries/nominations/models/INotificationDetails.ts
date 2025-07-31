import { IAllNominationDetails } from "./IAllNominationDetails";

export interface INotificationDetails {
  emailTitle?:string;
  emailTo: string;
  emailCC?:string;
  emailSub: string;
  emailBody:string;
  IsEnabled:boolean;
  allNominationData?:IAllNominationDetails;
}
