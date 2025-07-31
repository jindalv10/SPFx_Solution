
import { Presence as IUserPresence } from "@microsoft/microsoft-graph-types";

export enum Tables {
  PDNOMINATIONLIST = "PD Nominations",  
  NOMINATIONLOCALADMINLIST = "Nomination Details By Local Admin",
  NOMINATIONQCLIST = "Nomination Details By QC",
  NOMINATIONREVIEWERLIST = "Nomination Reviewers",
  NOMINATIONNOTIFICATIONLIST = "Nomination Notifications",
  NOMINATIONATTATTACHMENTSDOC = "Nomination Attachments",
}

export interface IFieldList {
  name: string;
  props: { FieldTypeKind: number, choices?: string[], richText?: boolean, default?: boolean | string };
}

export const QUESTIONLISTFields: IFieldList[] = [
  //  { name: "Question", props: { "FieldTypeKind": 2 } },
  { name: "ToolTip", props: { "FieldTypeKind": 3, "richText": false } },
  { name: "QuestionType", props: { "FieldTypeKind": 6, "choices": ["Yes/No", "Text"] } },
  { name: "Order", props: { "FieldTypeKind": 9 } },
  { name: "Enabled", props: { "FieldTypeKind": 8, "default": true } }
];

export const SELFCHECKINLISTFields: IFieldList[] = [
  { name: "CheckInOffice", props: { "FieldTypeKind": 2 } },
  { name: "Employee", props: { "FieldTypeKind": 20 } },
  { name: "Questions", props: { "FieldTypeKind": 3, "richText": false } }
];

export const COVIDCHECKINLISTFields: IFieldList[] = [
  { name: "CheckInOffice", props: { "FieldTypeKind": 2 } },
  { name: "Employee", props: { "FieldTypeKind": 20 } },
  { name: "Questions", props: { "FieldTypeKind": 3, "richText": false } },
  { name: "CheckIn", props: { "FieldTypeKind": 4 } },
  { name: "CheckInBy", props: { "FieldTypeKind": 20 } },
  { name: "Guest", props: { "FieldTypeKind": 2 } },
  { name: "SubmittedOn", props: { "FieldTypeKind": 4 } },
];


}