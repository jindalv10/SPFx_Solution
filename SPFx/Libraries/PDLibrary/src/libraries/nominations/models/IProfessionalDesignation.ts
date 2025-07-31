import { IDiscipline } from "./IDiscipline";

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
    discipline: number;
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
