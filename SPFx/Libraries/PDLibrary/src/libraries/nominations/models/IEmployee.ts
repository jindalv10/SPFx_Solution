export interface IEmployee {
  id: number;
  financeUserId?: number;
  email: string;
  pdDisciplines?: IDiscipline[];
  professionalDesignation?: IProfessionalDesignation;
}

export interface IDiscipline {
  id: number;
  name: string;
  friendlyName: string;
  abbreviation: string;
  created: Date;
  modified: Date;
  color: string;
}

export interface IProfessionalDesignation {
  id: number;
  code: string;
  title: string;
  created: Date;
  modified: Date;
  rank: number;
  createdBy?: string;
  modifiedBy?: string;
}

export interface IProfessionalDesignationDetailed {
    id: number;
    financeUserId: number;
    designationId: number;
    pdSubategoryId: number;
    discipline: IDiscipline;
    grantedOn: Date;
    removedOn: Date;
    level: any;
    restrictions: any;
    restrictionDate: any;
    professionalDesignation: string;
    subCategory: string;
    code: string;
    abbreviation: string;
    friendlyName: string;
    created: Date;
    modified: Date;
    createdBy: string;
    modifiedBy: string;

    isDeleted?: boolean;
    isNew?: boolean;
    isEdited?: boolean;
    isValid?: boolean;
}

export interface IEmployeeUpdateProperties {
    financeUserId: number;
    orientaionDate?: Date;
    advancedOrientation?: Date;
    satelliteOfficeId?: number;
    relatedBoardMember?: number;
    officerTitleID?: number;
    boardTitleId?: number;
    committeeAssignments?: { id: number, financeUserId: number, committeeId: number, isDelete: boolean }[];
    proficientLanguages?: {
      id: number,
      financeUserId: number,
      proficientLaguageId: string,
      isDelete: boolean
    }[];
    professionalDesignations: {
        id: number;
        financeUserId: number;
        designationId: number;
        pdSubategoryId: number;
        disciplineId: number;
        grantedOn: Date;
        removedOn: Date;
        level: any;
        isDelete: boolean;
    }[];
    shareholders?: {
        id: number;
        financeUserId: number;
        shareholderLegalId: number;
        shareholderId: number;
        date: Date;
        isDelete?: boolean;
    }[];
}
