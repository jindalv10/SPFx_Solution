export enum AllRoles {
    EP_NOMINATOR = "EP",
    NOMINATOR = "Nominator",
    LA = "Local Admin",
    QC = "QC",
    PTPAC_CHAIR = "Ptpac Chair",
    PTPAC_REVIEWER = "Ptpac Reviewer",
    LEGAL= "Legal"
}
export enum PdStatus {
    RP = "Recognized Professional",
    AP = "Approved Professional",
    LSA = "Limited Signature Authority",
    PSA = "Provisional Signature Authority",
    SA = "Signature Authority",
    QR = "Qualified Reviewer"
}
export enum QueryType {
    GETITEM = "GET_ITEM",
    GETFILES = "GET_FILES",
    UPDATE = "UPDATE",
    DELETE = "DELETE",
    ADD = "ADD",
    GETFIELD = "GET_FIELD"

}
export enum NominationStatus {
    DraftByNominator = "Draft By Nominator",
    Deleted = "Deleted",
    SubmittedByLocalAdmin = "Submitted By Local Admin",
    SubmittedByNominator = "Submitted by Nominator",
    PendingWithLocalAdmin = "Pending With Local Admin",
    PendingWithQC = "Pending With QC",
    PendingWithQCAndLegal = "Pending With QC and Legal",
    PendingWithPTPACChair = "Pending With PTPAC Chair",
    PendingWithPTPACReviewer = "Pending With PTPAC Reviewer",
    UnderQCReview ="Under QC Review",
    SubmittedByLegalPendingWithQc = "Submitted By Legal Pending With QC",
    SubmittedByQCPendingWithLegal = "Submitted By QC Pending With Legal",
    ApproveCompleted = "Approve-Completed",
    WithdrawnCompleted = "Withdraw - Completed",
    RequireAdditionalDetails = "Waiting For Additional Details",
    WithdrawRPStatus = "Withdraw - Reverse RP Status",
    Completed = "Completed"
}

export enum QCReviewStatus {
    QCDraft = "Draft By QC",
    Withdraw = "Withdraw - Reverse RP Status",
    RequireAdditionalDetails = "Waiting For Additional Details",
    SentToPTPAC = "Waiting For PTPAC Review",
    SentToSCForVote = "Waiting For SC Vote",
    SubmittedByQC = "Submitted By QC",
    GrantAccess = "Granted Access To Additional Reviewer",
}

export enum ReviewStatus {
  SubmittedByPTPACReviewer = "PTPAC review completed",
  SubmittedByPTPACChair = "PTPAC review completed"
}

export enum LegalStatus {
    SubmittedByLegal = "Submitted By Legal"
}

export interface IConstants {

    readonly ENV: string;
    readonly PowerAutomateFlowUrl?: string;
    readonly LOG_SOURCE: string;
    readonly SP_LIST_NAMES: {
        MasterNominationList: string,
        LANominationList: string,
        NominationDocumentLibraryName: string;
        QCNominationList: string,
        NominationReviewersList: string,
        PtpacNominationList: string,
        LegalNominationList: string,
        NotificationList: string,
        ReferencesList:string

    };
    readonly SP_TEST_USER: string;
    readonly SEND_EMAIL: boolean;
    readonly BREAK_ATTACHMENT_PERMISSION: boolean;

    readonly SP_Content_Type: {
        DocSetContentType: string;
    };
    readonly SP_Group: {
        LegalGroup: string;
    };
    readonly ENDPOINTS: {
        BaseAzureApiUrl: string,
        NomineeDetailsFetch: string,
        DisciplinesFetch: string,
        ProficientLanguageFetch: string,
        AadClientResourceIdentifier: string,
        ProfDesignationsFetch: string,
        PDSubcategoriesFetch: string,
        EmployeeDetailsUpdate:string,
        EmployeeProfessionalDesignationsByFinanceUserId:string,
        EmployeeInformationByFinanceUserId:string
    };
}
