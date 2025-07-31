import { INominationLibrary } from "./INominationLibrary";
import { IAllNominationDetails } from "../../models/IAllNominationDetails";
import { PnpAttachmentsRequest, PnpBatchReuqest } from "../../models/PnpBatchReuqest";
import { IUserDetails } from "../../models/IUserDetails";
import { AllRoles, NominationStatus, PdStatus, QCReviewStatus, QueryType, ReviewStatus } from "../../models/IConstants";
import { IIntakeNomination } from "../../models/IIntakeNomination";
import { INominationDetailsByLA } from "../../models/INominationDetailsByLA";
import { Utility } from "../startup/Utility";
import { Mapper } from "../startup/Mapper";
import { ISpUser } from "../../models/ISpUser";
import { INominationDetailsByQC } from "../../models/INominationDetailsByQC";
import { INominationDetailsByLegal } from "../../models/INominationDetailsByLegal";
import SPService from "../SPService";
import { INominationDetailsByPTPAC } from "../../models/INominationDetailsByPTPAC";
import { IReferences } from "../../models/IReferences";


export default class NominationLibrary extends SPService implements INominationLibrary {

    constructor(context: any) {
        super(context);
    }
    private getAllRequests(nominationId: number, nominee: ISpUser, currentUser: IUserDetails): PnpBatchReuqest[] {
        const isNominator = Utility.ciEquals(currentUser.role, AllRoles.NOMINATOR);
        const isLocalAdmin = Utility.ciEquals(currentUser.role, AllRoles.LA);
        const isQualityCoordinator = Utility.ciEquals(currentUser.role, AllRoles.QC);
        const isPTPACChair = Utility.ciEquals(currentUser.role, AllRoles.PTPAC_CHAIR);
        const isPTPACReviewer = Utility.ciEquals(currentUser.role, AllRoles.PTPAC_REVIEWER);

        let allRequest: PnpBatchReuqest[] = [
            {
                type: QueryType.GETITEM,
                list: this.Constants.SP_LIST_NAMES.MasterNominationList,
                id: nominationId,
                expand: "NomineeName,EPNominator,Author,References",
                select: "*,NomineeName/Title,NomineeName/EMail, EPNominator/Title,EPNominator/EMail,Author/Title,Author/EMail,References/Title,References/EMail"
            }
        ];

        if (currentUser) {
            if (isNominator || isLocalAdmin || isQualityCoordinator || isPTPACChair || isPTPACReviewer) {
                let laRequestObject: PnpBatchReuqest = {
                    type: QueryType.GETITEM,
                    list: this.Constants.SP_LIST_NAMES.LANominationList,
                    filter: "NomintaionId eq " + nominationId,
                    expand: "Nomintaion,Assignee",
                    select: "*,Nomintaion/Title,Assignee/Title,Assignee/EMail"
                };
                allRequest.push(laRequestObject);

                let referencesRequestObject: PnpBatchReuqest = {
                  type: QueryType.GETITEM,
                  list: this.Constants.SP_LIST_NAMES.ReferencesList,
                  filter: "NomintaionId eq " + nominationId,
                  expand: "Nomintaion,References",
                  select: "*,Nomintaion/Title,References/Title,References/EMail, ReferencesTrackStatus"
                };
                allRequest.push(referencesRequestObject);

            }
            if (isQualityCoordinator || isPTPACChair || isPTPACReviewer) {
                let qcRequestObject: PnpBatchReuqest = {
                    type: QueryType.GETITEM,
                    list: this.Constants.SP_LIST_NAMES.QCNominationList,
                    filter: "NomintaionId eq " + nominationId,
                    expand: "Nomintaion,Reviewer,AdditionalReviewer",
                    select: "*,Nomintaion/Title,AdditionalReviewer/Title,AdditionalReviewer/EMail,Reviewer/Title,Reviewer/EMail"
                };
                allRequest.push(qcRequestObject);

                let ptpacRequestObject: PnpBatchReuqest = {
                  type: QueryType.GETITEM,
                  list: this.Constants.SP_LIST_NAMES.PtpacNominationList,
                  filter: "NomintaionId eq " + nominationId,
                  expand: "Nomintaion,Reviewer,PTPACChair",
                  select: "*,Nomintaion/Title,Reviewer/Title,Reviewer/EMail,PTPACChair/Title,PTPACChair/EMail"
              };
              allRequest.push(ptpacRequestObject);
            }

            if (isQualityCoordinator) {
                let legalRequestObject: PnpBatchReuqest = {
                    type: QueryType.GETITEM,
                    list: this.Constants.SP_LIST_NAMES.LegalNominationList,
                    filter: "NomintaionId eq " + nominationId,
                    expand: "Nomintaion,Reviewer",
                    select: "*,Nomintaion/Title,Reviewer/Title,Reviewer/EMail"
                };
                allRequest.push(legalRequestObject);
            }
            if (isNominator || isQualityCoordinator || isPTPACChair || isPTPACReviewer) {
                let NominationAttachmentReqObject: PnpBatchReuqest = {
                    type: QueryType.GETFILES,
                    list: this.Constants.SP_LIST_NAMES.NominationDocumentLibraryName,
                    docSet: Utility.getDocSetName(nominee, nominationId),
                    folder: AllRoles.NOMINATOR
                };
                allRequest.push(NominationAttachmentReqObject);
            }
            if (isQualityCoordinator || isPTPACChair || isPTPACReviewer) {
                let attachmentReqObject: PnpBatchReuqest = {
                    type: QueryType.GETFILES,
                    list: this.Constants.SP_LIST_NAMES.NominationDocumentLibraryName,
                    docSet: Utility.getDocSetName(nominee, nominationId),
                    folder: AllRoles.QC
                };
                allRequest.push(attachmentReqObject);
            }
            if (isQualityCoordinator || isPTPACChair || isPTPACReviewer) {
              let attachmentReqObject: PnpBatchReuqest = {
                  type: QueryType.GETFILES,
                  list: this.Constants.SP_LIST_NAMES.NominationDocumentLibraryName,
                  docSet: Utility.getDocSetName(nominee, nominationId),
                  folder: AllRoles.PTPAC_CHAIR
              };
              allRequest.push(attachmentReqObject);
          }
          if (isQualityCoordinator || isPTPACChair || isPTPACReviewer) {
            let attachmentReqObject: PnpBatchReuqest = {
                type: QueryType.GETFILES,
                list: this.Constants.SP_LIST_NAMES.NominationDocumentLibraryName,
                docSet: Utility.getDocSetName(nominee, nominationId),
                folder: AllRoles.PTPAC_REVIEWER
            };
            allRequest.push(attachmentReqObject);
          }

        }
        return allRequest;
    }
    public async getNominationDetails(nominationId: number, nominee: ISpUser, currentUser: IUserDetails): Promise<IAllNominationDetails> {
        let _nominationDetails: IAllNominationDetails = {
            intakeNomination: null, nominationDetailsByLA: null, nominationAttachments: [], nominationReferences: [],
            nominationDetailsByLegal: null, nominationDetailsByPTPAC: null, nominationDetailsByQC: null
        };

        const allRequest = this.getAllRequests(nominationId, nominee, currentUser);
        await this.spBatchGet(allRequest).then((responses) => {
            console.info(responses);
            let intakeDetails, laDetails, qcDetails, legalDetails,ptpacDetails;
            if (responses) {
                intakeDetails = responses[0];
                let laRequest = allRequest.filter(x => x.list === this.Constants.SP_LIST_NAMES.LANominationList);
                let referencesRequest = allRequest.filter(x => x.list === this.Constants.SP_LIST_NAMES.ReferencesList);
                let legalRequest = allRequest.filter(x => x.list === this.Constants.SP_LIST_NAMES.LegalNominationList);
                let qcRequest = allRequest.filter(x => x.list === this.Constants.SP_LIST_NAMES.QCNominationList);
                let ptpacRequest = allRequest.filter(x => x.list === this.Constants.SP_LIST_NAMES.PtpacNominationList);
                let filesRequest = allRequest.filter(x => x.list === this.Constants.SP_LIST_NAMES.NominationDocumentLibraryName);

                if (laRequest && laRequest.length > 0) {
                    let laData = responses[allRequest.indexOf(laRequest[0])];
                    if (laData && laData.length > 0) {
                        laDetails = laData[0];
                    }
                }
                if (legalRequest && legalRequest.length > 0) {
                    let legalData = responses[allRequest.indexOf(legalRequest[0])];
                    if (legalData && legalData.length > 0) {
                        legalDetails = legalData[0];
                    }
                }
                if (qcRequest && qcRequest.length > 0) {
                    let qcData = responses[allRequest.indexOf(qcRequest[0])];
                    if (qcData && qcData.length > 0) {
                        qcDetails = qcData[0];
                    }
                }
                if (ptpacRequest && ptpacRequest.length > 0) {
                  let ptpacData = responses[allRequest.indexOf(ptpacRequest[0])];
                  if (ptpacData && ptpacData.length > 0) {
                    ptpacDetails = ptpacData[0];
                  }
                }
                if (filesRequest && filesRequest.length > 0) {
                    filesRequest.forEach(element => {
                        let allFilesResult = responses[allRequest.indexOf(element)];
                        if (allFilesResult) {
                            allFilesResult.forEach(nominationFilesResult => {
                                _nominationDetails.nominationAttachments.push(Mapper.mapAttachmentDetails(nominationFilesResult, element.folder));
                            });
                        }
                    });
                }
                if (referencesRequest && referencesRequest.length > 0 ) {
                  referencesRequest.forEach(element => {
                    let allReferencesResult = responses[allRequest.indexOf(element)];
                    if (allReferencesResult && allReferencesResult.length > 0) {
                      allReferencesResult.forEach(nominationReferencesResult => {
                            _nominationDetails.nominationReferences.push(Mapper.mapReferenceDetails(nominationReferencesResult, nominationReferencesResult.ReferencesId));
                      });
                    }
                    else if (intakeDetails && intakeDetails.References && intakeDetails.References.length > 0 )
                    {
                      intakeDetails.References.forEach((nominationReferencesResult, i) => {
                        _nominationDetails.nominationReferences.push(Mapper.mapReferenceDetails(nominationReferencesResult, intakeDetails.ReferencesId[i]));
                      });
                    }
                  });
                }


                if (intakeDetails)
                    _nominationDetails.intakeNomination = Mapper.mapIntakeNominationDetails(intakeDetails);
                if (laDetails)
                    _nominationDetails.nominationDetailsByLA = Mapper.mapLADetails(laDetails);
                if (legalDetails)
                    _nominationDetails.nominationDetailsByLegal = Mapper.mapLegalDetails(legalDetails);
                if (qcDetails)
                    _nominationDetails.nominationDetailsByQC = Mapper.mapQcDetails(qcDetails);
                if (ptpacDetails)
                    _nominationDetails.nominationDetailsByPTPAC = Mapper.mapPTPACDetails(ptpacDetails);


            }//if response end

        }).catch((e) => {
            console.log(e);
        });

        return Promise.resolve(_nominationDetails);
    }



    private buildStatusUpdateObject(intakeDetails: IIntakeNomination, status: string): PnpBatchReuqest {
        let intakeRequest: PnpBatchReuqest;
        if (intakeDetails && intakeDetails.id != 0) {
            let nominationId = intakeDetails.id;
            let _nominationDetails =
            {
                NominationStatus: status,
                IsProductPerson: intakeDetails.isProductPerson,
                TrackCandidateNominated: intakeDetails.trackCandidateNominated,
                BillingCode: intakeDetails.billingCode,
                ReferencesId: intakeDetails.references && intakeDetails.references.length > 0 ? {
                  results: intakeDetails.references.map((element) => { return element.id; }),
                } : {results:[0]},
            };
            intakeRequest = {
                type: nominationId == 0 ? QueryType.ADD : QueryType.UPDATE,
                list: this.Constants.SP_LIST_NAMES.MasterNominationList,
                data: _nominationDetails,
                id: nominationId
            };
        }
        return intakeRequest;
    }
    private getNominationStatus(nominationDetails: IAllNominationDetails, currentUser: IUserDetails): string {
      const isNominator = Utility.ciEquals(currentUser.role, AllRoles.NOMINATOR);
      const isLocalAdmin = Utility.ciEquals(currentUser.role, AllRoles.LA);
      const isQualityCoordinator = Utility.ciEquals(currentUser.role, AllRoles.QC);
      const isPTPACChair = Utility.ciEquals(currentUser.role, AllRoles.PTPAC_CHAIR);
      const isPTPACReviewer = Utility.ciEquals(currentUser.role, AllRoles.PTPAC_REVIEWER);
      let status = null;
      if (nominationDetails.intakeNomination) {
          let intakeDetails = nominationDetails.intakeNomination;
          if (isNominator) {
              status = intakeDetails.nominationStatus;
              if (intakeDetails.nominationStatus == NominationStatus.SubmittedByNominator) {
                  status = NominationStatus.PendingWithLocalAdmin;
                  if (intakeDetails.nomineeDesignation && intakeDetails.isStatusGrantedAfter2016 && intakeDetails.pdStatus !== PdStatus.RP) {
                    status = NominationStatus.PendingWithQC;
                  }
                  if (intakeDetails.isStatusGrantedAfter2016 && intakeDetails.pdStatus === PdStatus.RP) {
                    status = NominationStatus.ApproveCompleted;
                  }
              }
          }
          if (isLocalAdmin) {
              if (intakeDetails.nominationStatus == NominationStatus.SubmittedByLocalAdmin) {
                status = NominationStatus.PendingWithQC;
                if (intakeDetails.pdStatus === PdStatus.RP) {
                  status = NominationStatus.ApproveCompleted;
                }
            }
          }
          if (isQualityCoordinator) {
              let qcDetails = nominationDetails.nominationDetailsByQC;
              if (qcDetails.qcStatus == QCReviewStatus.Withdraw) {
                      status = NominationStatus.WithdrawnCompleted;
              }
              else if (qcDetails.qcStatus == QCReviewStatus.SubmittedByQC) {
                status = NominationStatus.ApproveCompleted;
              }
              else if (qcDetails.qcStatus == QCReviewStatus.SentToPTPAC) {
                status = NominationStatus.PendingWithPTPACChair;
              }
              else if (qcDetails.qcStatus == QCReviewStatus.SentToSCForVote || qcDetails.qcStatus == QCReviewStatus.QCDraft) {
                status = intakeDetails.nominationStatus ;
          }

          }
          if (isPTPACChair) {
            let qcDetails = nominationDetails.nominationDetailsByQC;

            if (qcDetails.qcStatus == QCReviewStatus.SentToPTPAC || qcDetails.qcStatus == QCReviewStatus.QCDraft) {
              status = NominationStatus.PendingWithPTPACReviewer;
            }
            else if (qcDetails.qcStatus == ReviewStatus.SubmittedByPTPACChair || qcDetails.qcStatus == QCReviewStatus.QCDraft) {
              status = NominationStatus.PendingWithQC;
            }

          }
          if (isPTPACReviewer) {
            let qcDetails = nominationDetails.nominationDetailsByQC;

            if (qcDetails.qcStatus == ReviewStatus.SubmittedByPTPACReviewer || qcDetails.qcStatus == QCReviewStatus.QCDraft) {
              status = NominationStatus.PendingWithPTPACChair;
            }

          }

      }
      return status;

    }
    private buildIntakeDetailsPostObject(intakeDetails: IIntakeNomination, nominationStatus: string): PnpBatchReuqest {
        let intakeRequest: PnpBatchReuqest;
        if (intakeDetails) {
            let nominationId = intakeDetails.id && intakeDetails.id != 0 ? intakeDetails.id : 0;
            let _nominationDetails = null;
            _nominationDetails =
            {
                Title: intakeDetails.title ? intakeDetails.title : intakeDetails.nominee && intakeDetails.nominee.title,
                DraftDate: intakeDetails.draftDate,
                IntakeNotes: intakeDetails.intakeNotes,
                IsProductPerson: intakeDetails.isProductPerson,
                IsStatusGrantedAfter2016: intakeDetails.isStatusGrantedAfter2016,
                NominationStatus: nominationStatus,
                EPNominatorId: intakeDetails.epNominators && intakeDetails.epNominators.length > 0 && {
                    results: intakeDetails.epNominators.map((element) => { return element.id; }),

                },
                NomineeDiscipline: intakeDetails.nomineeDiscipline,
                NomineeNameId: intakeDetails.nominee && intakeDetails.nominee.id,
                NomineeOffice: intakeDetails.nomineeOffice,
                NomineePractice: intakeDetails.nomineePractice,
                PDDiscipline: intakeDetails.pdDiscipline,
                PDStatus: intakeDetails.pdStatus,
                PDSubcategory: intakeDetails.pdSubcategory && intakeDetails.pdSubcategory.join(';'),
                ProficientLanguage: intakeDetails.proficientLanguage && intakeDetails.proficientLanguage.join(';'),
                RPCertification: intakeDetails.rpCertification,
                SubmissionDate: intakeDetails.submissionDate,
                ReSubmissionDate: intakeDetails.reSubmissionDate,
                FinanceUserID: intakeDetails.financeUserID,
                NomineeDesignation: intakeDetails.nomineeDesignation,
                TrackCandidateNominated: intakeDetails.trackCandidateNominated,
                ReferencesId: intakeDetails.references && intakeDetails.references.length > 0 ? {
                  results: intakeDetails.references.map((element) => { return element.id; }),
                } : {results:[0]},
                BillingCode: intakeDetails.billingCode,
                //GrantDate: intakeDetails.grantDate
            };
            intakeRequest = {
                type: nominationId == 0 ? QueryType.ADD : intakeDetails.nominationStatus == NominationStatus.Deleted ? QueryType.DELETE : QueryType.UPDATE,
                list: this.Constants.SP_LIST_NAMES.MasterNominationList,
                data: _nominationDetails,
                id: nominationId
            };
        }
        return intakeRequest;
    }

    private buildReferencesPostObject(referencesDetails: IReferences, nominationId: number, nominationStatus: string): PnpBatchReuqest {
      let referencesRequest: PnpBatchReuqest;
      if (nominationStatus && referencesDetails) {
          let referencesDetailsId = referencesDetails.id && referencesDetails.id != 0 ? referencesDetails.id : 0;
          let _nominationReferenceDetails =
          {
              Title: referencesDetails.referencesUser ? referencesDetails.referencesUser.title : '',
              ReferencesId: referencesDetails.referencesUser ? referencesDetails.referencesUser.id : '',
              ReferencesTrackStatus: referencesDetails.referencesTrackVal,
              NomintaionId: nominationId

          };
          referencesRequest = {
              type: referencesDetails.referencesUser == null || nominationStatus == NominationStatus.Deleted ? QueryType.DELETE : referencesDetailsId == 0 ? QueryType.ADD : QueryType.UPDATE,
              list: this.Constants.SP_LIST_NAMES.ReferencesList,
              data: _nominationReferenceDetails,
              id: referencesDetailsId
          };
      }
      return referencesRequest;
   }
    private buildLocalAdminPostObject(localAdminDetails: INominationDetailsByLA, nominationId: number, nominationStatus: string): PnpBatchReuqest {
        let laRequest: PnpBatchReuqest;
        if (nominationStatus && localAdminDetails) {
            let localAdminDetailsId = localAdminDetails.id && localAdminDetails.id != 0 ? localAdminDetails.id : 0;
            let _nominationDetails =
            {
                Title: localAdminDetails.title,
                AssigneeId: localAdminDetails.assignee && localAdminDetails.assignee.id,
                IsEmployeeAgreementSigned: localAdminDetails.isEmployeeAgreementSigned,
                IsEmployeeNumberUpdated: localAdminDetails.isEmployeeNumberUpdated,
                ReviewNotes: localAdminDetails.reviewNotes,
                ReviewDate: localAdminDetails.reviewDate,
                WithdrawCompletionDate: localAdminDetails.withdrawCompletionDate,
                EmployeeNumberReversedDate: localAdminDetails.employeeNumberReversedDate,
                IsEmployeeNumberReversed: localAdminDetails.isEmployeeNumberReversed,
                NomintaionId: nominationId

            };
            laRequest = {
                type: localAdminDetailsId == 0 ? QueryType.ADD : !localAdminDetails.assignee || nominationStatus == NominationStatus.Deleted ? QueryType.DELETE : QueryType.UPDATE,
                list: this.Constants.SP_LIST_NAMES.LANominationList,
                data: _nominationDetails,
                id: localAdminDetailsId
            };
        }
        return laRequest;
    }
    private buildLegalPostObject(legalDetails: INominationDetailsByLegal, nominationId: number): PnpBatchReuqest {
        let laRequest: PnpBatchReuqest;
        if (legalDetails) {
            let legalAdminDetailsId = legalDetails.id && legalDetails.id != 0 ? legalDetails.id : 0;
            let _nominationDetails =
            {
                Title: legalDetails.title,
                AssigneeId: legalDetails.reviewer && legalDetails.reviewer.id,
                IsEmpAgreementSignedByCEO: legalDetails.isEmpAgreementSignedByCEO,
                IsSavedOnLocalDrive: legalDetails.isSavedOnLocalDrive,
                ReviewDate: legalDetails.reviewDate,
                NomintaionId: nominationId

            };
            laRequest = {
                type: legalAdminDetailsId == 0 ? QueryType.ADD : QueryType.UPDATE,
                list: this.Constants.SP_LIST_NAMES.LegalNominationList,
                data: _nominationDetails,
                id: legalAdminDetailsId
            };
        }
        return laRequest;
    }
    private buildQCPostObject(qcDetails: INominationDetailsByQC, nominationId: number): PnpBatchReuqest {
        let qcRequest: PnpBatchReuqest;
        if (qcDetails) {
            let qcDetailsId = qcDetails.id && qcDetails.id != 0 ? qcDetails.id : 0;

            let _nominationDetails =
            {
                Title: nominationId + "-" + qcDetails.qcStatus,
                ReviewerId: qcDetails.reviewer && qcDetails.reviewer.id,
                ReviewNotes: qcDetails.reviewNotes,
                ReviewDate: qcDetails.reviewDate,
                NomintaionId: nominationId,
                QCStatus: qcDetails.qcStatus,
                SentForMoreDetails: qcDetails.sentForMoreDetails,
                SentToPTPACDate: qcDetails.sentToPTPACDate,
                WithdrawnDate: qcDetails.withdrawnDate,
                SentToScDate: qcDetails.sentToScDate,
                DraftDate: qcDetails.draftDate,
                // Additional Reviewer Date
                ReviewerAssignmentDate: qcDetails.reviewerAssignmentDate,
                AdditionalReviewerId: qcDetails.additionalReviewer && qcDetails.additionalReviewer.id,
                // Status Granted
                Granted: qcDetails.granted,
                NominationEndDate: qcDetails.endDate,
                NotifcationRecipient: qcDetails.notificationRecipient,
                NominationPasses: qcDetails.nominationPasses,
                ReferencesPassed: qcDetails.referencesPassed,
                QARPassed: qcDetails.qarPassed,
                AddPracticeDirector: qcDetails.addPracticeDirector ? true : false,
                AnyoneElse: qcDetails.anyoneElse ? qcDetails.anyoneElse : null,

            };
            qcRequest = {
                type: qcDetailsId == 0 ? QueryType.ADD : QueryType.UPDATE,
                list: this.Constants.SP_LIST_NAMES.QCNominationList,
                data: _nominationDetails,
                id: qcDetailsId
            };
        }
        return qcRequest;
    }

    private buildPTPACPostObject(ptpacReviewDetails: INominationDetailsByPTPAC, nominationId: number):PnpBatchReuqest {
      let ptpacRequest: PnpBatchReuqest;
      if (ptpacReviewDetails) {
          let ptpacDetailsId = ptpacReviewDetails.id && ptpacReviewDetails.id != 0 ? ptpacReviewDetails.id : 0;

          let _nominationDetails =
          {
              Title: nominationId + "-" + "PTAPC Details",
              ReviewerId: ptpacReviewDetails.reviewer && ptpacReviewDetails.reviewer.id,
              PTPACChairId: ptpacReviewDetails.ptpacChair && ptpacReviewDetails.ptpacChair.id,
              Recommendation:ptpacReviewDetails.recommendation,
              ReviewDueDate: ptpacReviewDetails.reviewDueDate,
              ReviewDate: ptpacReviewDetails.reviewDate,
              RecommendationSentDate: ptpacReviewDetails.recommendationSentDate,
              ReviewerAssignmentDate: ptpacReviewDetails.reviewerAssignmentDate,
              NomintaionId: nominationId,
              ptpacChairComments: ptpacReviewDetails.ptpacChairComments,
              internalReviewDueDate: ptpacReviewDetails.internalReviewDueDate && new Date(ptpacReviewDetails.internalReviewDueDate)
          };
          ptpacRequest = {
              type: ptpacDetailsId == 0 ? QueryType.ADD : QueryType.UPDATE,
              list: this.Constants.SP_LIST_NAMES.PtpacNominationList,
              data: _nominationDetails,
              id: ptpacDetailsId
          };
      }
     return  ptpacRequest;
    }

    public async deleteFile(nominationDetails: IAllNominationDetails, currentUser: IUserDetails, subFolderName: string, fileName: string) : Promise<boolean>{
      if (nominationDetails && currentUser) {
        let nominationId = nominationDetails.intakeNomination.id ? nominationDetails.intakeNomination.id : 0;
        let docSetName = null;
        docSetName = Utility.getDocSetName(nominationDetails.intakeNomination.nominee, nominationId);
        await this.delFile(this.Constants.SP_LIST_NAMES.NominationDocumentLibraryName, docSetName, subFolderName, fileName).then((response) => {
          console.info(response);
        }).catch((e) => {
            console.log(e);
            return Promise.reject(false);
        });
        return Promise.resolve(true);
      }
    }


    public async saveNominationDetails(nominationDetails: IAllNominationDetails, currentUser: IUserDetails, assignPermission: ISpUser[], groupName: string[]): Promise<boolean> {
        const isNominator = Utility.ciEquals(currentUser.role, AllRoles.NOMINATOR);
        const isLocalAdmin = Utility.ciEquals(currentUser.role, AllRoles.LA);
        const isQualityCoordinator = Utility.ciEquals(currentUser.role, AllRoles.QC);
        const isPTPACChair = Utility.ciEquals(currentUser.role, AllRoles.PTPAC_CHAIR);
        const isPTPACReviewer = Utility.ciEquals(currentUser.role, AllRoles.PTPAC_REVIEWER);
        let responseData = null;

        if (nominationDetails && currentUser) {
            if (nominationDetails.intakeNomination) {
                let nominationId = nominationDetails.intakeNomination.id ? nominationDetails.intakeNomination.id : 0;
                let allRequest: PnpBatchReuqest[] = [];
                let allAttachmentRequest: PnpAttachmentsRequest = null;
                let docSetName = null;
                let intakeRequest: PnpBatchReuqest = null;
                let nominationStatus = this.getNominationStatus(nominationDetails, currentUser);
                if (isNominator) {
                    docSetName = Utility.getDocSetName(nominationDetails.intakeNomination.nominee, nominationId);
                    if (nominationId == 0) {
                        intakeRequest = this.buildIntakeDetailsPostObject(nominationDetails.intakeNomination, nominationStatus);
                        await this.spPost(intakeRequest).then(async (id) => {
                            nominationId = id;
                            docSetName = Utility.getDocSetName(nominationDetails.intakeNomination.nominee, nominationId);
                            const documentSetDetails = await this.createDocumentSet(this.Constants.SP_LIST_NAMES.NominationDocumentLibraryName, docSetName, this.Constants.SP_Content_Type.DocSetContentType, Utility.getRolesForFolder());
                            console.log(documentSetDetails);
                        }).catch((e) => {
                            console.error("Cannot add nomination request");
                            return Promise.resolve(false);
                        });
                    }
                    else {
                        if (nominationStatus == NominationStatus.Deleted)
                            await this.deleteDocumentSet(this.Constants.SP_LIST_NAMES.NominationDocumentLibraryName, docSetName);
                        intakeRequest = this.buildIntakeDetailsPostObject(nominationDetails.intakeNomination, nominationStatus);
                        allRequest.push(intakeRequest);
                    }
                    if (nominationDetails.nominationDetailsByLA && nominationDetails.nominationDetailsByLA.assignee) {
                        let laRequest = this.buildLocalAdminPostObject(nominationDetails.nominationDetailsByLA, nominationId, nominationStatus);
                        allRequest.push(laRequest);
                    }
                    if (nominationDetails.intakeNomination && nominationDetails.intakeNomination.references) {
                        nominationDetails.intakeNomination.references.map((element) => {
                          allRequest.push(this.buildReferencesPostObject(element, nominationId, nominationStatus));
                        });
                    }
                }
                if (nominationDetails.nominationAttachments && nominationId != 0) {
                    docSetName = Utility.getDocSetName(nominationDetails.intakeNomination.nominee, nominationId);
                    allAttachmentRequest = {
                        list: this.Constants.SP_LIST_NAMES.NominationDocumentLibraryName,
                        docSet: docSetName,
                        attachments: nominationDetails.nominationAttachments,
                    };
                }
                if (nominationDetails.nominationDetailsByLA && nominationId != 0 && isLocalAdmin) {
                    let laRequest = this.buildLocalAdminPostObject(nominationDetails.nominationDetailsByLA, nominationId, nominationStatus);
                    allRequest.push(laRequest);
                }
                if (nominationDetails.nominationDetailsByQC && nominationId != 0 && isQualityCoordinator || isPTPACChair || isPTPACReviewer) {
                    let qcRequest = this.buildQCPostObject(nominationDetails.nominationDetailsByQC, nominationId);
                    let ptpacRequest = this.buildPTPACPostObject(nominationDetails.nominationDetailsByPTPAC, nominationId);
                    let newIntakeRequest  = this.buildIntakeDetailsPostObject(nominationDetails.intakeNomination, nominationStatus);
                    let laRequest = this.buildLocalAdminPostObject(nominationDetails.nominationDetailsByLA, nominationId, nominationStatus);

                    allRequest.push(qcRequest);
                    if(ptpacRequest)
                      allRequest.push(ptpacRequest);
                    if(newIntakeRequest )
                      allRequest.push(newIntakeRequest );
                    if(laRequest)
                      allRequest.push(laRequest);
                }
                if (nominationDetails.nominationDetailsByLegal && nominationId != 0) {
                    let laRequest = this.buildLegalPostObject(nominationDetails.nominationDetailsByLegal, nominationId);
                    allRequest.push(laRequest);
                }
                if (!isNominator) {
                    intakeRequest = this.buildStatusUpdateObject(nominationDetails.intakeNomination, nominationStatus);
                    allRequest.push(intakeRequest);
                }
                if (nominationId != 0 && nominationDetails.nominationReferences && nominationDetails.nominationReferences.length > 0 && isQualityCoordinator || isNominator) {
                  nominationDetails.nominationReferences.map((element) => {
                    let referenceRequest = this.buildReferencesPostObject(element, nominationId, nominationStatus);
                    allRequest.push(referenceRequest);
                  });
                }
                /*
                if (!isLocalAdmin && !groupName) {
                  try {
                    await this.assignItemLevelPermission(this.Constants.SP_LIST_NAMES.NominationDocumentLibraryName, assignPermission, docSetName, groupName);
                  } catch (permissionError) {
                    console.log("Permission assignment failed:", permissionError);
                  }
                }
                */
                if ((allRequest && allRequest.length > 0) || allAttachmentRequest) {
                  try {
                    responseData = await this.spBatchPostAll(allRequest, allAttachmentRequest);
                    responseData = ({responseVal: responseData, attachmentDocumentSetName: docSetName });
                    return Promise.resolve(responseData);
                  } catch (batchError) {
                    console.log("Batch processing failed:", batchError);
                    return Promise.reject(false);
                  }
              }
            }
        }
    }
}
