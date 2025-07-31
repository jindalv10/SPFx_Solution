import { DialogType, ICommandBarStyles, ITheme, createTheme, IToggleStyles, IButtonStyles, mergeStyles, IStackTokens } from 'office-ui-fabric-react';
import {
    IColumn
} from '@fluentui/react';
import { INominationListViewItem } from 'pd-nomination-library';
import CommonMethods from '../models/CommonMethods';
import { NominationStatus } from 'pd-nomination-library';

const ThemeColorsFromWindow: any = (window as any).__themeState__.theme;
const siteTheme: ITheme = createTheme({
    palette: ThemeColorsFromWindow
});

export const getSPFormatDate = (date: Date) => {
  if (date) {
      //return date.getFullYear() + "-" + (date.getMonth() + 1) + "-" + date.getDate() + "T07:00:00Z";
      //}
      var current = new Date();
      return new Date(date.getFullYear() + "/" + (date.getMonth() + 1) + "/" + date.getDate() + " " + current.getHours() + ":" + current.getMinutes() + ":" + current.getSeconds()).toISOString();
  }
  else
      return null;
};

export enum PanelPosition {
    Left,
    Right
}
export const commandBarStyles: Partial<ICommandBarStyles> = {
    root: {
        backgroundColor: siteTheme.palette.themePrimary,
    },
    primarySet: {
        backgroundColor: siteTheme.palette.themePrimary
    },
    secondarySet: {
        backgroundColor: siteTheme.palette.themePrimary
    }

};
export const margin = '0 20px 20px 0';
export const controlWrapperClass = mergeStyles({
    display: 'flex',
    flexWrap: 'wrap',
});
export const toggleStyles: Partial<IToggleStyles> = {
    root: { margin: margin },
    label: { marginLeft: 10 },
};
const addItemButtonStyles: Partial<IButtonStyles> = { root: { margin: margin } };


export const STATUS = {
    DELETE_SUCCESS: 'DELETE_SUCCESS',
    DELETE_ERROR: 'DELETE_SUCCESS',
    SAVE_ERROR: 'SAVE_ERROR',
    SAVE_SUCCESS: 'SAVE_SUCCESS',
    SUBMIT_ERROR: 'SUBMIT_ERROR',
    SUBMIT_SUCCESS: 'SUBMIT_SUCCESS',
    DEFAULT: 'DEFAULT'
};
export const QCBUTTONSACTIONS = {
    REQUESTS_MORE_DETAILS: 'REQUESTS_MORE_DETAILS',
    WITHDRAW_NOMINATION: 'Withdraw Nomination',
    SEND_SC_FOR_VOTE: 'Send to SC for Vote',
    GRANT_STATUS: 'Grant Status',
    GRANT_ACCESS_TO_SOMEONE: 'GRANT_ACCESS_TO_SOMEONE',
    REQUEST_PTPAC_REVIEW: 'Request PTPAC Review',
    SEND_EMAIL: 'Send Email',
};

export const PTPACBUTTONSACTIONS = {
  ASSIGN_A_PTPAC_REVIEWER: 'Assign a PTPAC Reviewer',
  SEND_TO_QC: 'Send To QC',
  SEND_TO_PTPAC_CHAIR: 'Send to PTPAC Chair'
};

export const GENERICBUTTONSACTIONS = {
  SUBMIT: 'Submit',
  SAVE:"Save",
  CANCEL: 'Cancel',
  DELETE: "Delete",
  SEND:"Send",
  SAVEANDCLOSE:"Save & Close",
};

export const GROUPNAME = {
  EPGroup: 'Professional Designation Nomination EPs Group',
  NominationOwners:"Professional Designation Nomination Owners",
  QCGroup: 'Professional Designation Nomination QC Group',
  NominationVisitors: "Professional Designation Nomination Visitors",
  SCGroup: 'Professional Designation SC Group',

};

export const Messages = {
    DeleteFailedPrefix: "Some error occurred while deleting request!",
    DeleteSuccessPrefix: "Request has been deleted successfully!",
    SaveFailedPrefix: "Some error occurred while saving request!",
    SaveSuccessPrefix: "Request has been saved successfully!",
    SubmitFailedPrefix: "Some error occurred while submitting request!",
    SubmitSuccessPrefix: "Request has been submitted successfully!",
    DeletConfirmationMessage: "Are you sure you want to delete this request? Once deleted cannot be recovered."
};
export const dialogContentProps = {
    type: DialogType.normal,
    title: Messages.DeletConfirmationMessage,
};


export const dialogModalProps = {
    isBlocking: true,
    styles: { main: { maxWidth: 450 } },
};
export const stackTokens: IStackTokens = { childrenGap: 40 };

export const NominationListColumns: IColumn[] = [{
    key: 'Nominee',
    name: 'Nominee',
    fieldName: 'Nominee',
    minWidth: 150,
    maxWidth: 150,
    isSortedDescending: false,
    // isSorted: false,
    isResizable: false,
    data: String,

},
{
    key: 'InternalStatus',
    name: 'Internal Status',
    fieldName: 'InternalStatus',
    minWidth: 150,
    maxWidth: 150,
    isResizable: true,
    isSorted: true,
    isSortedDescending: false,
    onRender: (item: INominationListViewItem) => {
        let status: string =  null;
        switch(item["InternalStatus"])
        {
            case NominationStatus.DraftByNominator: {
                status = "Draft";
                break;
            }
            case NominationStatus.SubmittedByNominator: 
            case NominationStatus.PendingWithQC:
            case NominationStatus.PendingWithLocalAdmin: 
            case NominationStatus.PendingWithPTPACChair:
            case NominationStatus.PendingWithPTPACReviewer: {
                status = "In Progress";
                break;
            }
            case NominationStatus.ApproveCompleted:{
                status = "Completed";
                break;
            }
            case NominationStatus.WithdrawnCompleted: {
              status = "Withdrawn";
              break;
          }

        }
        return status;
    },
    // sortAscendingAriaLabel: 'Sorted A to Z',
    // sortDescendingAriaLabel: 'Sorted Z to A',
    data: String
},
{
    key: 'PDStatus',
    name: 'PD Status',
    fieldName: 'PDStatus',
    minWidth: 150,
    maxWidth: 150,
    isResizable: true,
    // isSorted: false,
    isSortedDescending: false,
    data: String
},
{
    key: 'PDDiscipline',
    name: 'PD Discipline',
    fieldName: 'PDDiscipline',
    minWidth: 100,
    maxWidth: 100,
    isResizable: true,
    // isSorted: false,
    isSortedDescending: false,
    data: String
},
{
    key: 'Nominator',
    name: 'Nominator',
    fieldName: 'Nominator',
    minWidth: 150,
    maxWidth: 150,
    isResizable: true,
    // isSorted: false,
    // isSortedDescending: false,
    data: String
},
{
    key: 'EPnominator',
    name: 'EP nominator',
    fieldName: 'EPnominator',
    minWidth: 150,
    maxWidth: 150,
    isResizable: true,
    // isSorted: false,
    isSortedDescending: false,
    data: Date
},
{
    key: 'Submitted',
    name: 'Submitted',
    fieldName: 'Submitted',
    minWidth: 100,
    maxWidth: 100,
    isResizable: true,
    // isSorted: true,
    isSortedDescending: true,
    data: String
},
];

export const OtherDetailsListColumns: IColumn[] =
      [{
        key: 'Nominee',
        name: 'Nominee',
        fieldName: 'Nominee',
        minWidth: 150,
        maxWidth: 150,
        isResizable: false,
        data: String
      },
      {
        key: 'InternalStatus',
        name: 'Internal Status',
        fieldName: 'InternalStatus',
        minWidth: 150,
        maxWidth: 150,
        isResizable: true,
        isSorted: false,
        isSortedDescending: false,
        data: String
      },
      {
        key: 'PDStatus',
        name: 'PD Status',
        fieldName: 'PDStatus',
        minWidth: 150,
        maxWidth: 150,
        isResizable: true,
        isSorted: false,
        isSortedDescending: false,
        data: String
      },
      {
        key: 'PDDiscipline',
        name: 'PD Discipline',
        fieldName: 'PDDiscipline',
        minWidth: 100,
        maxWidth: 100,
        isResizable: true,
        isSorted: false,
        isSortedDescending: false,
        data: String
      },
      {
        key: 'Nominator',
        name: 'Nominator',
        fieldName: 'Nominator',
        minWidth: 150,
        maxWidth: 150,
        isResizable: true,
        isSorted: false,
        isSortedDescending: false,
        data: String
      },
      {
        key: 'EPnominator(s)',
        name: 'EP nominator',
        fieldName: 'EPnominator',
        minWidth: 150,
        maxWidth: 150,
        isResizable: true,
        isSorted: false,
        isSortedDescending: false,
        data: String
      },
      {
        key: 'Submitted',
        name: 'Submitted',
        fieldName: 'Submitted',
        minWidth: 100,
        maxWidth: 100,
        isResizable: true,
        isSorted: false,
        isSortedDescending: false,
        data: String
      },
];

export const QCDetailsListColumns: IColumn[] =
      [{
        key: 'Nominee',
        name: 'Nominee',
        fieldName: 'Nominee',
        minWidth: 150,
        maxWidth: 150,
        isResizable: false,
        data: String
      },
      {
        key: 'InternalStatus',
        name: 'Internal Status',
        fieldName: 'InternalStatus',
        minWidth: 150,
        maxWidth: 150,
        isResizable: true,
        isSorted: false,
        isSortedDescending: false,
        data: String
      },
      {
        key: 'SendSCforVoteDate',
        name: 'Send SC for Vote Date',
        fieldName: 'SendSCforVoteDate',
        minWidth: 150,
        maxWidth: 150,
        isResizable: true,
        isSorted: false,
        isSortedDescending: false,
        data: String
      },
      {
        key: 'NominationPasses',
        name: 'Nomination Passes',
        fieldName: 'NominationPasses',
        minWidth: 150,
        maxWidth: 150,
        isResizable: true,
        isSorted: false,
        isSortedDescending: false,
        data: String
      },
      {
        key: 'ReferencesPassed',
        name: 'References Passed',
        fieldName: 'ReferencesPassed',
        minWidth: 150,
        maxWidth: 150,
        isResizable: true,
        isSorted: false,
        isSortedDescending: false,
        data: String
      },
      {
        key: 'QARPassed',
        name: 'QAR Passed',
        fieldName: 'QARPassed',
        minWidth: 150,
        maxWidth: 150,
        isResizable: true,
        isSorted: false,
        isSortedDescending: false,
        data: String
      },
      {
        key: 'PDStatus',
        name: 'PD Status',
        fieldName: 'PDStatus',
        minWidth: 150,
        maxWidth: 150,
        isResizable: true,
        isSorted: false,
        isSortedDescending: false,
        data: String
      },
      {
        key: 'Subcategory',
        name: 'Subcategory',
        fieldName: 'Subcategory',
        minWidth: 150,
        maxWidth: 150,
        isResizable: true,
        isSorted: false,
        isSortedDescending: false,
        data: String
      },
      {
        key: 'Nominator',
        name: 'Nominator',
        fieldName: 'Nominator',
        minWidth: 150,
        maxWidth: 150,
        isResizable: true,
        isSorted: false,
        isSortedDescending: false,
        data: String
      },
      {
        key: 'Submitted',
        name: 'Submitted',
        fieldName: 'Submitted',
        minWidth: 150,
        maxWidth: 150,
        isResizable: true,
        isSorted: false,
        isSortedDescending: false,
        data: String
      },
      {
        key: 'PDDiscipline',
        name: 'PD Discipline',
        fieldName: 'PDDiscipline',
        minWidth: 150,
        maxWidth: 150,
        isResizable: true,
        isSorted: false,
        isSortedDescending: false,
        data: String
      }
];

export const PTPACChairDetailsListColumns: IColumn[] =
      [{
        key: 'Nominee',
        name: 'Nominee',
        fieldName: 'Nominee',
        minWidth: 150,
        maxWidth: 150,
        isResizable: false,
        data: String
      },
      {
        key: 'InternalStatus',
        name: 'Internal Status',
        fieldName: 'InternalStatus',
        minWidth: 150,
        maxWidth: 150,
        isResizable: true,
        isSorted: false,
        isSortedDescending: false,
        data: String
      },
      {
        key: 'PDStatus',
        name: 'PD Status',
        fieldName: 'PDStatus',
        minWidth: 150,
        maxWidth: 150,
        isResizable: true,
        isSorted: false,
        isSortedDescending: false,
        data: String
      },
      {
        key: 'PDDiscipline',
        name: 'PD Discipline',
        fieldName: 'PDDiscipline',
        minWidth: 100,
        maxWidth: 100,
        isResizable: true,
        isSorted: false,
        isSortedDescending: false,
        data: String
      },
      {
        key: 'PTPACDueDate',
        name: 'PTPAC Due Date',
        fieldName: 'PTPACDueDate',
        minWidth: 150,
        maxWidth: 150,
        isResizable: true,
        isSorted: false,
        isSortedDescending: false,
        data: Date
      },
      {
        key: 'PTPACInternalDueDate',
        name: 'PTPAC Internal Due Date',
        fieldName: 'PTPACInternalDueDate',
        minWidth: 200,
        maxWidth: 200,
        isResizable: false,
        isSorted: false,
        isSortedDescending: false,
        data: Date
      },
      {
        key: 'PTPACReviewer',
        name: 'PTPAC Reviewer',
        fieldName: 'PTPACReviewer',
        minWidth: 100,
        maxWidth: 100,
        isResizable: false,
        isSorted: false,
        isSortedDescending: false,
        data: String
      },
];

export const PTPACReviewerDetailsListColumns: IColumn[] =
      [{
        key: 'Nominee',
        name: 'Nominee',
        fieldName: 'Nominee',
        minWidth: 200,
        maxWidth: 200,
        isResizable: false,
        data: String
      },
      {
        key: 'InternalStatus',
        name: 'Internal Status',
        fieldName: 'InternalStatus',
        minWidth: 200,
        maxWidth: 200,
        isResizable: true,
        isSorted: false,
        isSortedDescending: false,
        data: String
      },
      {
        key: 'PTPACInternalDueDate',
        name: 'PTPAC Internal Due Date',
        fieldName: 'PTPACInternalDueDate',
        minWidth: 200,
        maxWidth: 200,
        isResizable: false,
        isSorted: true,
        isSortedDescending: true,
        data: Date
      },
      {
        key: 'PDStatus',
        name: 'PD Status',
        fieldName: 'PDStatus',
        minWidth: 250,
        maxWidth: 250,
        isResizable: false,
        isSorted: false,
        isSortedDescending: false,
        data: String
      },
      {
        key: 'PDDiscipline',
        name: 'PD Discipline',
        fieldName: 'PDDiscipline',
        minWidth: 250,
        maxWidth: 250,
        isResizable: false,
        isSorted: false,
        isSortedDescending: false,
        data: String
      },
      {
        key: 'PTPACDueDate',
        name: 'PTPAC Due Date',
        fieldName: 'PTPACDueDate',
        minWidth: 200,
        maxWidth: 200,
        isResizable: false,
        isSorted: false,
        isSortedDescending: false,
        data: Date
      }
];


