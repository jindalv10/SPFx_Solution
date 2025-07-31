import { IAllNominationDetails } from "../../models/IAllNominationDetails";
import { IAttachment } from "../../models/IAttachment";
import { IIntakeNomination } from "../../models/IIntakeNomination";
import { INominationDetailsByLA } from "../../models/INominationDetailsByLA";
import { INominationDetailsByLegal } from "../../models/INominationDetailsByLegal";
import { INominationDetailsByPTPAC } from "../../models/INominationDetailsByPTPAC";
import { INominationDetailsByQC } from "../../models/INominationDetailsByQC";
import { INominationListViewItem } from "../../models/INominationListViewItem";
import { INominationReviewer } from "../../models/INominationReviewer";
import { INotificationDetails } from "../../models/INotificationDetails";
import { IReferences } from "../../models/IReferences";
import { Utility } from "./Utility";

export class Mapper {

    public static mapNominationListDetails(item): INominationListViewItem {
        let nomination: INominationListViewItem = null;
        if (item) {

            nomination = {
                id: item.ID,
                status: item.NominationStatus,
                nominee: item.NomineeName ? item.NomineeName.length > 0 ? { id: item.NomineeName[0].id, title: item.NomineeName[0].title } : null : null,
                epNominators: item.EPNominator && item.EPNominator.map((element) => { return { id: element.id, title: element.title }; }),
                pdDiscipline: item.PDDiscipline,
                pdStatus: item.PDStatus,
                nominator: item.Author ? item.Author.length > 0 ? { id: item.Author[0].id, title: item.Author[0].title } : null : null,
                submitted: item.SubmissionDate,
                sendSCforVoteDate: item.hasOwnProperty('baseKey') && item["baseKey"]["sentToScDate"] ? item["baseKey"]["sentToScDate"]  : null,
                nominationPasses:  item.hasOwnProperty('baseKey') && item["baseKey"]["nominationPasses"] ? item["baseKey"]["nominationPasses"] : false,
                referencesPassed:  item.hasOwnProperty('baseKey') && item["baseKey"]["referencesPassed"]  ? item["baseKey"]["referencesPassed"] : false,
                qarPassed:   item.hasOwnProperty('baseKey') && item["baseKey"]["qarPassed"] ? item["baseKey"]["qarPassed"] : false,
                PTPACDueDate:   item.hasOwnProperty('baseKey') && item["baseKey"]["reviewDueDate"] ? item["baseKey"]["reviewDueDate"] : null,
                PTPACInternalDueDate:   item.hasOwnProperty('baseKey') && item["baseKey"]["internalReviewDueDate"] ? item["baseKey"]["internalReviewDueDate"]  : null,
                PTPACReviewer:  item.hasOwnProperty('baseKey') && item["baseKey"]["reviewer"]  ? { id:item["baseKey"]["reviewer"].id, title: item["baseKey"]["reviewer"].title } : null,
                Subcategory: item.PDSubcategory && item.PDSubcategory.split(';').map((element) => { return element; }),


            };
        }
        return nomination;
    }

    public static mapNotificationDetailsList(item, allNomationDataInfo: IAllNominationDetails): INotificationDetails {
      let notification: INotificationDetails = null;
      if (item) {


        const emailSubjectAfterReplace = (emailSubject: string) => {
          const emailSubjectArray: Array<string> = emailSubject.match(/[^{{]+(?=\}})/g);
          let replaceEmailSubject: string = item.emailSub;
          if(emailSubjectArray){
            emailSubjectArray.forEach(element => {
              const nestedVal = Utility.findDeepNestedObject(allNomationDataInfo, element);
              replaceEmailSubject = replaceEmailSubject.replace("{{" + element  + "}}", nestedVal);
            });
          }
          return replaceEmailSubject;
        };

        const emailBodyAfterReplace = (emailBody: string) => {
          const emailBodyArray: Array<string> = emailBody.match(/[^{{]+(?=\}})/g);
          let replaceEmailBody: string = item.emailBody;
          if(emailBodyArray){
            emailBodyArray.forEach(element => {
              const nestedVal = Utility.findDeepNestedObject(allNomationDataInfo, element);
              replaceEmailBody = replaceEmailBody.replace("{{" + element  + "}}", nestedVal);
            });
          }
          return replaceEmailBody;
        };

        notification = {
              emailTitle: item.emailTitle,
              emailTo: item.emailTo,
              emailCC: item.emailCC ? item.emailCC : null,
              emailSub: item.emailSub ? emailSubjectAfterReplace(item.emailSub) : null,
              emailBody: item.emailBody ? emailBodyAfterReplace(item.emailBody) : null,
              IsEnabled: item.IsEnabled.toLowerCase() === 'true' ? true : false,
          };
      }
      return notification;
    }

    public static mapQCReviewerDisciplineDetails(reviewerDetails): INominationReviewer {
      let nominationReviewer: INominationReviewer = {
          AuthorizedQC: reviewerDetails.AuthorizedQC && reviewerDetails.AuthorizedQC.map((element, i) => { return { id: reviewerDetails.AuthorizedQC[i].id, title: reviewerDetails.AuthorizedQC[i].title, email: reviewerDetails.AuthorizedQC[i].email }; }),
          AuthorizedPTPAC: reviewerDetails.AuthorizedPTPAC && reviewerDetails.AuthorizedPTPAC.map((element, i) => { return { id: reviewerDetails.AuthorizedPTPAC[i].id, title: reviewerDetails.AuthorizedPTPAC[i].title, email: reviewerDetails.AuthorizedPTPAC[i].email }; }),
          PDDiscipline: reviewerDetails.PDDiscipline
      };
      return nominationReviewer;
  }

    public static mapIntakeNominationDetails(intakeDetails): IIntakeNomination {

        let intakeNomination: IIntakeNomination = {
            title: intakeDetails.Title,
            id: intakeDetails.Id,
            draftDate: intakeDetails.DraftDate,
            intakeNotes: intakeDetails.IntakeNotes,
            isProductPerson: intakeDetails.IsProductPerson,
            isStatusGrantedAfter2016: intakeDetails.IsStatusGrantedAfter2016,
            nominationStatus: intakeDetails.NominationStatus,
            epNominators: intakeDetails.EPNominator && intakeDetails.EPNominator.map((element, i) => { return { id: intakeDetails.EPNominatorId[i], title: element.Title, email: element.EMail }; }),
            nomineeDiscipline: intakeDetails.NomineeDiscipline,
            nomineeDesignation: intakeDetails.NomineeDesignation,
            nominee: intakeDetails.NomineeName && { id: intakeDetails.NomineeNameId, title: intakeDetails.NomineeName.Title, email: intakeDetails.NomineeName.EMail },
            nomineeOffice: intakeDetails.NomineeOffice,
            nomineePractice: intakeDetails.NomineePractice,
            pdDiscipline: intakeDetails.PDDiscipline,
            pdStatus: intakeDetails.PDStatus,
            pdSubcategory: intakeDetails.PDSubcategory ? intakeDetails.PDSubcategory.split(";") : null,
            proficientLanguage: intakeDetails.ProficientLanguage ? intakeDetails.ProficientLanguage.split(";") : null,
            rpCertification: intakeDetails.RPCertification,
            submissionDate: intakeDetails.SubmissionDate,
            reSubmissionDate: intakeDetails.ReSubmissionDate,
            nominator: intakeDetails.Author ? { id: intakeDetails.AuthorId, title: intakeDetails.Author.Title, email: intakeDetails.Author.EMail } : null,
            financeUserID: intakeDetails.FinanceUserID ? intakeDetails.FinanceUserID : null,
            trackCandidateNominated: intakeDetails.TrackCandidateNominated ? intakeDetails.TrackCandidateNominated : null,
            references : intakeDetails.References && intakeDetails.References.map((element, i) => { return { id: intakeDetails.ReferencesId[i], title: element.Title, email: element.EMail }; }),
            billingCode: intakeDetails.BillingCode ? intakeDetails.BillingCode : null,
        };
        return intakeNomination;
    }
    public static mapLADetails(laDetails): INominationDetailsByLA {

        let nominationDetailsByLA: INominationDetailsByLA = {
            title: laDetails.Title,
            assignee: laDetails.Assignee && { id: laDetails.AssigneeId, title: laDetails.Assignee.Title, email: laDetails.Assignee.EMail },
            isEmployeeAgreementSigned: laDetails.IsEmployeeAgreementSigned,
            id: laDetails.Id,
            employeeNumberReversedDate: laDetails.EmployeeNumberReversedDate,
            isEmployeeNumberReversed: laDetails.IsEmployeeNumberReversed,
            isEmployeeNumberUpdated: laDetails.IsEmployeeNumberUpdated,
            nominationId: laDetails.NomintaionId,
            reviewDate: laDetails.ReviewDate,
            reviewNotes: laDetails.ReviewNotes,
            withdrawCompletionDate: laDetails.WithdrawCompletionDate
        };
        return nominationDetailsByLA;
    }
    public static mapLegalDetails(legalDetails): INominationDetailsByLegal {

        let nominationDetailsByLegal: INominationDetailsByLegal = {
            title: legalDetails.Title,
            reviewer: legalDetails.Reviewer && { id: legalDetails.ReviewerId, title: legalDetails.Reviewer.Title, email: legalDetails.Reviewer.EMail },
            isEmpAgreementSignedByCEO: legalDetails.IsEmpAgreementSignedByCEO,
            id: legalDetails.Id,
            isSavedOnLocalDrive: legalDetails.IsSavedOnLocalDrive,
            nominationId: legalDetails.NomintaionId,
            reviewDate: legalDetails.ReviewDate,

        };
        return nominationDetailsByLegal;
    }
    public static mapQcDetails(qcDetails): INominationDetailsByQC {

        let nominationDetailsByQC: INominationDetailsByQC = {
            id: qcDetails.Id,
            reviewer: qcDetails.Reviewer && { id: qcDetails.ReviewerId, title: qcDetails.Reviewer.Title, email: qcDetails.Reviewer.EMail },
            reviewNotes: qcDetails.ReviewNotes,
            reviewDate: qcDetails.ReviewDate,
            sentToScDate: qcDetails.SentToScDate,
            draftDate: qcDetails.DraftDate,
            withdrawnDate: qcDetails.WithdrawnDate,
            sentToPTPACDate:  qcDetails.SentToPTPACDate,
            notificationRecipient: qcDetails.NotifcationRecipient && qcDetails.NotifcationRecipient.join(';'),
            nominationId: qcDetails.NomintaionId,
            sentForMoreDetails: qcDetails.SentForMoreDetails,
            qcStatus: qcDetails.QCStatus,
            reviewerAssignmentDate: qcDetails.ReviewerAssignmentDate,
            additionalReviewer: qcDetails.AdditionalReviewer && { id: qcDetails.AdditionalReviewerId, title: qcDetails.AdditionalReviewer.Title },
            anyoneElse: qcDetails.anyoneElse,
            addPracticeDirector: qcDetails.addPracticeDirector,
            nominationPasses: qcDetails.NominationPasses,
            referencesPassed: qcDetails.ReferencesPassed,
            qarPassed: qcDetails.QARPassed

        };
        return nominationDetailsByQC;
    }
    public static mapPTPACDetails(ptpacDetails): INominationDetailsByPTPAC {

      let nominationDetailsByPTPAC: INominationDetailsByPTPAC = {
          id: ptpacDetails.Id,
          reviewer: ptpacDetails.Reviewer && { id: ptpacDetails.ReviewerId, title: ptpacDetails.Reviewer.Title, email: ptpacDetails.Reviewer.EMail },
          ptpacChair: ptpacDetails.PTPACChair && { id: ptpacDetails.PTPACChairId, title: ptpacDetails.PTPACChair.Title, email: ptpacDetails.PTPACChair.EMail },
          recommendation: ptpacDetails.Recommendation,
          reviewDueDate: ptpacDetails.ReviewDueDate,
          reviewDate:ptpacDetails.ReviewDate,
          recommendationSentDate: ptpacDetails.RecommendationSentDate,
          reviewerAssignmentDate: ptpacDetails.ReviewerAssignmentDate,
          nominationId: ptpacDetails.NomintaionId,
          ptpacChairComments: ptpacDetails.ptpacChairComments,
          internalReviewDueDate: ptpacDetails.internalReviewDueDate,
      };
      return nominationDetailsByPTPAC;
  }
    public static mapAttachmentDetails(nominationFilesResult, folder): IAttachment {

        let nominationAttachments: IAttachment = {
            id: nominationFilesResult.ListItemAllFields.Id,
            attachmentName: nominationFilesResult.Name,
            attachmentUrl: nominationFilesResult.ServerRelativeUrl,
            attachmentType: nominationFilesResult.ListItemAllFields.AttachmentType,
            attachmentBy: folder
        };
        return nominationAttachments;
    }

    public static mapReferenceDetails(nominationReferencesResult, ReferencesId?: number ): IReferences {

      let nominationReferences: IReferences = {
          id:  nominationReferencesResult && nominationReferencesResult.Id ? nominationReferencesResult.Id : 0,
          referencesUser: {id: ReferencesId, title: nominationReferencesResult.Title ? nominationReferencesResult.Title : nominationReferencesResult.References.Title, email: nominationReferencesResult.EMail ? nominationReferencesResult.EMail : nominationReferencesResult.References.EMail },
          referencesTrackVal: nominationReferencesResult.ReferencesTrackStatus ? nominationReferencesResult.ReferencesTrackStatus : "Blank",
      };
      return nominationReferences;
  }



}
