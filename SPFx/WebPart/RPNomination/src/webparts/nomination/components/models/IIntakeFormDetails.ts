export enum FormName {
  "Intake",
  "LocalAdmin",
  "PTPAC",
  "QC",
  "GCSLead"
}

export enum FormField {
    id,
    nomineeName,
    epNominators,
    nomineeOffice,
    nomineePractice,
    nomineeDiscipline,
    isProductPerson,
    pdDiscipline,
    pdStatus,
    pdSubcategory,
    intakeNotes,
    nominationStatus,
    granted,
    nominationEndDate,
    rpCertification,
    submissionDate,
    draftDate,
    reSubmissionDate,
    proficientLanguage,
    isStatusGrantedAfter2016,
    attachments
  }
  
  export class FormError {
    public nomineeName: string = "";
    public epNominators: string = "";
    public nomineeOffice: string = "";
    public nomineePractice: string = "";
    public nomineeDiscipline: string = "";
    public isProductPerson: boolean = null;
    public pdDiscipline: string = "";
    public pdStatus: string = "";
    public pdSubcategory: string = "";
    public intakeNotes:string = "";
    public nominationStatus:string = "";
    public granted: Date = null;
    public nominationEndDate: Date = null;
    public rpCertification: boolean = null;
    public submissionDate:Date = null;
    public draftDate:Date = null;
    public reSubmissionDate: Date = null;
    public proficientLanguage: string;
    public isStatusGrantedAfter2016: boolean = null;
    public attachments: string = null;
  
  }