export interface IEmployeeUpdateProperties {
    financeUserId: number;
    orientaionDate: Date;
    advancedOrientation: Date;
    satelliteOfficeId: number;
    relatedBoardMember: number;
    officerTitleID: number;
    boardTitleId: number;
    committeeAssignments: { id: number, financeUserId: number, committeeId: number, isDelete: boolean }[];
    proficientLanguages: { id: number, financeUserId: number, proficientLaguageId: string, isDelete: boolean }[];
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
    shareholders: {
        id: number;
        financeUserId: number;
        shareholderLegalId: number;
        shareholderId: number;
        date: Date;
        isDelete?: boolean;
    }[];
}
