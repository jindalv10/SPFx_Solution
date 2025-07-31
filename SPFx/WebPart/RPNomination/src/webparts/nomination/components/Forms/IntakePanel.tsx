import * as React from 'react';
import *  as NominationLibraryComponent from "pd-nomination-library";
import styles from './Panel.module.scss';
import { DefaultButton, Dropdown, IDropdownOption, MessageBar, MessageBarButton, MessageBarType, Panel, PanelType, PrimaryButton, Stack, TextField, Toggle } from '@fluentui/react';
import { INomineeExist, INominationReviewer ,IAllNominationDetails, IAttachment, IIntakeNomination, IMasterDetails, INomineeDetails, IReferences } from 'pd-nomination-library';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IPeoplePickerUserItem, PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { AllRoles, NominationStatus, PdStatus, QCReviewStatus } from 'pd-nomination-library';
import { INominationListViewItem } from 'pd-nomination-library';
import { FileUploader } from '../control/FileUploader';
import autobind from 'autobind-decorator';
import SpinnerComponent from '../spinnerComponent/spinnerComponent';
import { dialogContentProps, dialogModalProps, GENERICBUTTONSACTIONS, GROUPNAME, Messages, PanelPosition, QCBUTTONSACTIONS, stackTokens, STATUS } from '../commonSettings/settings';
import { Dialog, DialogFooter, DialogType } from 'office-ui-fabric-react';
import { ConstantsConfig, IConstants, INITIAL_CANDIDATE_NOMINATED } from '../models/IUIConstants';
import { IEmployeeUpdateProperties } from 'pd-nomination-library';
import { IProfessionalDesignationDetailed } from 'pd-nomination-library';
import { SPHttpClient, IHttpClientOptions } from '@microsoft/sp-http';
import CommonMethods from '../models/CommonMethods';


export interface IIntakePanelProps {
    position?: PanelPosition;
    onDismiss?: () => void;
    context: WebPartContext;
    invokedItem: INominationListViewItem;
    isNewForm: boolean;
}


export interface IIntakePanelState {
    isOpen?: boolean;
    isFormStatus?: string;
    isSaveValid: boolean;
    isSubmitValid: boolean;
    isReferencesValid: {minRequired:number,isRequired:boolean, isHide:boolean};
    nomineeDetails?: INomineeDetails;
    detailsNominationReferences: IReferences[];
    itemDetails: IAllNominationDetails;
    intakeNomination: IIntakeNomination;
    masterListData: IMasterDetails;
    existingSubcategory:string[];
    loading: boolean;
    NominationFormAttachment: IAttachment[];
    NominationOtherAttachments: IAttachment[];
    attachmentType: string;
    files: Array<any>;
    isValidNominee: boolean;
    isValidNomineeEPNominator: string;
    actionStatus: string;
    isConfirmationDialogVisible: boolean;
    isMessageDialogVisible: boolean;
    isLARequired: boolean;
    nominationReviewersUsers: INominationReviewer[];
    pdNominationDetailed: IProfessionalDesignationDetailed[];
    grantedOn?: Date;

}
export default class IntakeFormPanel extends React.Component<IIntakePanelProps, IIntakePanelState> {

    public masterDetails: IMasterDetails;
    private NominationLibComponent = new NominationLibraryComponent.NominationLibrary(this.props.context);
    private NominationLoggerComponent = new NominationLibraryComponent.CustomLogger(this.props.context);

    private NominationListLibComponent = new NominationLibraryComponent.NominationListLibrary(this.props.context);
    private NominationLibMasters = new NominationLibraryComponent.IntakeNominationLibrary(this.props.context);
    private EmailNotification = new NominationLibraryComponent.NotificationList(this.props.context);
    protected Constants: IConstants = null;

    private intakeFormDetails = null;
    public constructor(props: IIntakePanelProps, state: IIntakePanelState) {
        super(props, state);

        this.state = {
            itemDetails: null,
            intakeNomination: null,
            isOpen: true,
            nomineeDetails: null,
            detailsNominationReferences: [],
            masterListData: null,
            existingSubcategory:[],
            loading: !this.props.isNewForm,
            NominationFormAttachment: [],
            NominationOtherAttachments: [],
            attachmentType: "Other",
            files: [],
            isSaveValid: false,
            isSubmitValid: false,
            isReferencesValid:null,
            isValidNominee: false,
            isValidNomineeEPNominator: null,
            actionStatus: null,
            isConfirmationDialogVisible: false,
            isMessageDialogVisible: false,
            isLARequired: false,
            nominationReviewersUsers: null,
            pdNominationDetailed: null,
            grantedOn: new Date(Date.now()),

        };
        this.Constants = ConstantsConfig.GetConstants();
    }


    private isValidationError(type: number) {
        const { intakeNomination, isLARequired, NominationFormAttachment, itemDetails, isValidNominee, isReferencesValid, detailsNominationReferences } = this.state;
        let isError = intakeNomination ? false : true;
        if (intakeNomination) {
            let laDetails = itemDetails ? itemDetails.nominationDetailsByLA : null;
            if (!isValidNominee && !intakeNomination.nominationStatus)
                isError = true;
            if (!intakeNomination.pdDiscipline && type == 1)
                isError = true;
            if (intakeNomination.pdDiscipline === "Employee Benefits" && !intakeNomination.pdSubcategory && type == 1)
                isError = true;
            else if (intakeNomination.pdDiscipline === "Employee Benefits" && intakeNomination.pdSubcategory && intakeNomination.pdSubcategory.length == 0 && type == 1)
                isError = true;
            if (!intakeNomination.pdStatus && type == 1)
                isError = true;
            if (!intakeNomination.epNominators && type == 1)
                isError = true;
            else if (intakeNomination.epNominators && intakeNomination.epNominators.length == 0 && type == 1) {
                isError = true;
            }
            if (!intakeNomination.proficientLanguage && type == 1)
                isError = true;
            if (intakeNomination && !intakeNomination.billingCode && intakeNomination.trackCandidateNominated && intakeNomination.pdStatus && intakeNomination.pdStatus.toUpperCase() === "SIGNATURE AUTHORITY" && intakeNomination.trackCandidateNominated === INITIAL_CANDIDATE_NOMINATED[2].key && type == 1)
                isError = true;    
            else if (intakeNomination.proficientLanguage && intakeNomination.proficientLanguage.length == 0 && type == 1) {
                isError = true;
            }
            if (isLARequired && !laDetails && type == 1) {
                isError = true;
            }
            else if (isLARequired && laDetails && !laDetails.assignee && type == 1) {
                isError = true;
            }
           
            // if (!intakeNomination && intakeNomination.intakeNotes && type == 1)
            //     isError = true;
            if (intakeNomination.pdStatus != PdStatus.RP && !NominationFormAttachment && type == 1)
                isError = true;
            else if (intakeNomination.pdStatus != PdStatus.RP && NominationFormAttachment && NominationFormAttachment.length == 0 && type == 1) {
                isError = true;
            }
            if (intakeNomination.pdStatus == PdStatus.RP && !intakeNomination.rpCertification && type == 1)
                isError = true;
            if (!intakeNomination.trackCandidateNominated  && type == 1)
                isError = true;    

            if(isReferencesValid && type == 1)
            {
                const notNullReferencesUserLength = detailsNominationReferences && detailsNominationReferences.filter(references => references.referencesUser !== null).length;
                if (detailsNominationReferences && detailsNominationReferences !== undefined && notNullReferencesUserLength < isReferencesValid.minRequired)
                    isError =  isReferencesValid.isRequired;
            }    

        }
        return isError;
    }



    private NominationTrackMinLimitAndRequired()
    {
        const { intakeNomination} = this.state;
        let validation = {minRequired: 0, isRequired:false, isHide:true};
        if(intakeNomination && intakeNomination.pdStatus && intakeNomination.trackCandidateNominated){
            if(intakeNomination.pdStatus.toUpperCase() === "APPROVED PROFESSIONAL" 
                && intakeNomination.trackCandidateNominated === INITIAL_CANDIDATE_NOMINATED[1].key 
                || intakeNomination.trackCandidateNominated === INITIAL_CANDIDATE_NOMINATED[2].key){
                validation = {minRequired: 2, isRequired:true, isHide: false};
            }
            if((intakeNomination.pdStatus.toUpperCase() === "SIGNATURE AUTHORITY" || intakeNomination.pdStatus.toUpperCase() === "QUALIFIED REVIEWER") 
                    && (intakeNomination.trackCandidateNominated === INITIAL_CANDIDATE_NOMINATED[0].key 
                        || intakeNomination.trackCandidateNominated === INITIAL_CANDIDATE_NOMINATED[1].key 
                        || intakeNomination.trackCandidateNominated === INITIAL_CANDIDATE_NOMINATED[2].key)){
                validation = {minRequired: 3, isRequired:true, isHide: false};
            }
            if(intakeNomination.pdStatus.toUpperCase() === "LIMITED SIGNATURE AUTHORITY" 
                    && (intakeNomination.trackCandidateNominated === INITIAL_CANDIDATE_NOMINATED[0].key 
                        || intakeNomination.trackCandidateNominated === INITIAL_CANDIDATE_NOMINATED[1].key 
                        || intakeNomination.trackCandidateNominated === INITIAL_CANDIDATE_NOMINATED[2].key)){
                validation = {minRequired: 0, isRequired:false, isHide: false};
            }
            if(intakeNomination.pdStatus.toUpperCase() === "RECOGNIZED PROFESSIONAL"){
                validation = {minRequired: 0, isRequired:false, isHide: true};
            }
             return validation;
        }
        return validation;
    }
    private SetValidState() {

        this.setState({
            isSaveValid: !this.isValidationError(0),
            isSubmitValid: !this.isValidationError(1),
            isReferencesValid: this.NominationTrackMinLimitAndRequired(),
            loading: false
        });
    }


    private _getNomineeInformation = async (items: any[]) => {

     try{
            let nomineeAPIData: INomineeDetails = null;
            let clearValue = false;
            const pickerValID = items && items.length > 0 ? items[0].id : ""; 
            const validNominee: INomineeExist = await this.NominationLibMasters.checkIfValidNominee(pickerValID);
            if (validNominee) {
                const itemData: IPeoplePickerUserItem[] = items;
                const originalEmail = itemData[0].secondaryText;
                const pickerReplaceVal = originalEmail 
                ? originalEmail.replace('millimandev.com', 'milliman.com').replace('millimantest.com', 'milliman.com') 
                : originalEmail;                //nomineeAPIData = await this._handleAsync(this.NominationLibMasters.getNomineeDetailsFromEmpDB("donna.boyle@milliman.com"));
                nomineeAPIData = await this._handleAsync(
                    this.NominationLibMasters.getNomineeDetailsFromEmpDB(pickerReplaceVal)
                );

                // Check if response is 204 (No Content), then retry with original test domain
                if (nomineeAPIData === null) {                    
                    const testDomainEmail = originalEmail?.replace('millimandev.com', 'milliman.com').replace('milliman.com', 'millimantest.com');
                    
                    nomineeAPIData = await this._handleAsync(
                        this.NominationLibMasters.getNomineeDetailsFromEmpDB(testDomainEmail)
                    );
                }
            }
            else {
                clearValue = true;
            }
            // this.setState({ loading: true }, async () => {


            if (nomineeAPIData) {
                this.setState((prevState) => ({
                    intakeNomination: {
                        ...prevState.intakeNomination, 
                        proficientLanguage: ["English"],
                        nominee: { title: items[0].hasOwnProperty("text") == true ? items[0].text: items[0].title  , id: items[0].id, email: items[0].secondaryText },
                        nomineeOffice: nomineeAPIData ? nomineeAPIData.office : null,
                        nomineeDesignation: nomineeAPIData ? nomineeAPIData.designation : null,
                        nomineePractice: nomineeAPIData ? nomineeAPIData.practice : null,
                        nomineeDiscipline: nomineeAPIData ? nomineeAPIData.discipline : null,
                        isStatusGrantedAfter2016: nomineeAPIData ? nomineeAPIData.isStatusGrantedAfter2016 : false,
                        financeUserID: nomineeAPIData.financeUserId,

                    },
                    nomineeDetails: nomineeAPIData,
                    loading: false,
                    isValidNominee: true,
                    //isValidNomineeEPNominator:validNominee.EPNominator,
                    isLARequired : nomineeAPIData && nomineeAPIData.isStatusGrantedAfter2016 ? !nomineeAPIData.isStatusGrantedAfter2016 : true
                    
                }), () => {
                    this.SetValidState();
                    this._handleAsync(this._getNomineeProfessionalDesignation());
                    this._checkIfNomineeStatusInProgressBasedOnDesignationAndDiscipline();

                });

                
            }
            if (clearValue) {

                this.setState((prevState) => ({
                    intakeNomination: {
                        ...prevState.intakeNomination,
                        nominee: items && items.length > 0 ? { title: items[0].text, id: items[0].id , email: items[0].secondaryText  } : null,
                        nomineeOffice: null,
                        nomineePractice: null,
                        nomineeDiscipline: null,
                        nomineeDesignation: null,
                        isStatusGrantedAfter2016: false,
                        financeUserID: null,

                    },
                    nomineeDetails: null,
                    loading: false,
                    isValidNominee: true,
                    //isValidNomineeEPNominator:validNominee.EPNominator,
                    isLARequired : false
                }), () => {
                    this.SetValidState();
                });
            }
        }
        catch (error) {
            this.NominationLoggerComponent.Error({
                WebPartName: "PD Nomination",
                ComponentName: "Intake Submission",
                MethodName: "_getNomineeInformation",
                Message: "Exception Occurred : " + error.message
            });
        }

    }

    private findExcludedIndexes(arr1, arr2) {
        const excludedIndexes = [];
      
        for (let i = 0; i < arr1.length; i++) {
          let found = false;
      
          for (let j = 0; j < arr2.length; j++) {
            if (this.isEqual(arr1[i], arr2[j])) {
              found = true;
              break;
            }
          }
      
          if (!found) {
            excludedIndexes.push(i);
          }
        }
      
        return excludedIndexes;
      }
      
      private isEqual(obj1, obj2) {
        // Implement your own logic to compare the objects
        // For example, compare each property of the objects
        // and return true if they are equal, otherwise false
        return obj1.referencesUser !== null && obj1.referencesUser.title === obj2.text;
      }

    
    private _getRefereeInformation = (items: any) => {
        const {detailsNominationReferences} = this.state;

        try {
            if (items.length > 0) {
                if(detailsNominationReferences.length < items.length)
                {
                    detailsNominationReferences.push({id: 0, referencesUser: { title: items[items.length-1].text, id: items[items.length-1].id, email: items[items.length-1].secondaryText }, referencesTrackVal: ""});
                }
                else if(detailsNominationReferences && detailsNominationReferences.length > items.length)
                {
                    let indexNumber: any = this.findExcludedIndexes(detailsNominationReferences, items);
                    indexNumber.map(i => detailsNominationReferences[i].referencesUser = null);

                }
                else if(detailsNominationReferences && detailsNominationReferences.length === items.length)
                {
                    const filteredArray = detailsNominationReferences.filter(obj => obj.referencesUser === null);
                    const result = filteredArray.length > 0 ? detailsNominationReferences.indexOf(filteredArray[0]) : 0;
                    detailsNominationReferences[result].referencesUser = { title: items[items.length-1].text, id: items[items.length-1].id, email: items[items.length-1].secondaryText };
                }
                this.setState((prevState) => ({
                    detailsNominationReferences: detailsNominationReferences,
                    loading: false
                }), () => {
                    this.SetValidState();
                });
            }
            
            else {
                this.setState((prevState) => ({
                    detailsNominationReferences: [],
                    itemDetails: {
                        ...prevState.itemDetails,
                        nominationReferences: [],
                    },
                    loading: false
                }), () => {
                    this.SetValidState();
                });
            }
        } catch (error) {
            this.NominationLoggerComponent.Error({
                WebPartName: "PD Nomination",
                ComponentName: "Intake Submission",
                MethodName: "_getRefereeInformation",
                Message: "Exception Occurred : " + error.message
            });
        }
    }



    private _getNominatorInformation = (items: any[]) => {

        try {
            const { itemDetails } = this.state;

            if (items.length > 0) {
                let laDetails: NominationLibraryComponent.INominationDetailsByLA = {
                    title: "LA Details",
                    id: this.state && this.state.itemDetails && this.state.itemDetails.nominationDetailsByLA && this.state.itemDetails.nominationDetailsByLA.id,
                    assignee: { title: items[0].text, id: items[0].id, email: items[0].secondaryText }
                };
                this.setState((prevState) => ({
                    itemDetails: {
                        ...prevState.itemDetails,
                        nominationDetailsByLA: laDetails,
                    },
                    loading: false
                }), () => {
                    this.SetValidState();
                });
            }
            else if (itemDetails && itemDetails.nominationDetailsByLA && itemDetails.nominationDetailsByLA.id) {
                this.setState((prevState) => ({
                    itemDetails: {
                        ...prevState.itemDetails,
                        nominationDetailsByLA: {
                            id: itemDetails.nominationDetailsByLA.id,
                            assignee: null,
                            title: itemDetails.nominationDetailsByLA.title,
                            
                        },
                    },
                    loading: false
                }), () => {
                    this.SetValidState();
                });
            }
            else {
                this.setState((prevState) => ({
                    itemDetails: {
                        ...prevState.itemDetails,
                        nominationDetailsByLA: null
                    },
                    loading: false
                }), () => {
                    this.SetValidState();
                });
            }
        } catch (error) {
            this.NominationLoggerComponent.Error({
                WebPartName: "PD Nomination",
                ComponentName: "Intake Submission",
                MethodName: "_getNominatorInformation",
                Message: "Exception Occurred : " + error.message
            });
        }

        
    }
    private _getEPNominatorInformation = (items: any) => {
       const {intakeNomination}  = this.state;
       let filterEPNominator = [];

        try {
            if (items.length > 0) {
                //const newItems = items.length > 1 ? intakeNomination.epNominators.map(item => ({ text: item.title, id: item.id, secondaryText: item.email })) : items;
                //const itemsIdNotNull = [...newItems, ...items.filter(a => a.id !== undefined)]'
                if(intakeNomination.epNominators)
                {
                    filterEPNominator =   intakeNomination.epNominators.filter(obj1 => items.find(obj2 => obj1.title === obj2.text));
                }
        
                const userExists = (arr, name) => {
                    const { length } = arr;
                    const id = length + 1;
                    const found = arr.some(el => el.title === name);
                    if (!found) arr.push({ title: items[items.length-1].text, id: items[items.length-1].id, email: items[items.length-1].secondaryText });
                    return arr;
                };
        
                filterEPNominator = userExists(filterEPNominator,items[items.length-1].text);

                this.setState((prevState) => ({
                    intakeNomination: {
                        ...prevState.intakeNomination,
                        epNominators: filterEPNominator

                    },
                    loading: false
                }), () => {
                    this.SetValidState();
                });
            }
            else {
                this.setState((prevState) => ({
                    intakeNomination: {
                        ...prevState.intakeNomination,
                        epNominators: []
                    },
                    loading: false
                }), () => {
                    this.SetValidState();
                });
            }
        } catch (error) {
            this.NominationLoggerComponent.Error({
                WebPartName: "PD Nomination",
                ComponentName: "Intake Submission",
                MethodName: "_getEPNominatorInformation",
                Message: "Exception Occurred : " + error.message
            });
        }
    }

    private async getUserInfo(email:string) {

     try   
        {
        const rootUrl = this.props.context.pageContext.site.absoluteUrl;
        let options = null;
        let spOptions: IHttpClientOptions = options ? options : {
            headers: {
                'Accept': 'application/json;odata=nometadata',
            }
        };
        return this.props.context.httpClient.get(
            rootUrl + "/_api/web/siteusers/getbyemail('" + email + "')",
            SPHttpClient.configurations.v1,
            spOptions
        )
        .then(
            (response: any) => {
                if (response.status >= 200 && response.status < 300) {
                    return response;
                } else {
                    return Promise.reject(JSON.stringify(response));
                }
            });
        } catch (error) {
            this.NominationLoggerComponent.Error({
                WebPartName: "PD Nomination",
                ComponentName: "Intake Submission",
                MethodName: "getUserInfo",
                Message: "Exception Occurred : " + error.message
            });
        }
    }
    

    public async componentDidMount() {
        try{
            this.initializeComponent();
            //https://millimandev.sharepoint.com/teams/DeveloperSite-MLT/SitePages/Nominee-Forms.aspx?Actor=Nominator&loadSPFX=true&debugManifestsFile=https%3A%2F%2Flocalhost%3A4321%2Ftemp%2Fmanifests.js
        
            let params = (new URL(window.location.href)).searchParams;
            // const formType = params.get('FormType'); // is the string "NewForm".
        
            const panel = params.get('Panel'); // is the string "NewForm".
            if(panel){
                const NomineeEmail = params.get('NomineeEmail'); // is the string "NewForm".
                this.getUserInfo(NomineeEmail).then((response) => { return response.json(); })
                .then(async (data) => {
                    if (data) {
                        console.log(data);
                        const items = [{ title: data.Title, id: data.Id, secondaryText: data.Email }];
                        this.setState((prevState) => ({
                            intakeNomination: {
                                ...prevState.intakeNomination, 
                                nominee: { title: data.Title, id: data.Id, secondaryText: data.Email  },
                            },
                            isValidNominee:true
                        }), () => {
                            this._getNomineeInformation(items);
                        });
                    }
                });    
            }
        } catch (error) {
            this.NominationLoggerComponent.Error({
                WebPartName: "PD Nomination",
                ComponentName: "Intake Submission",
                MethodName: "componentDidMount",
                Message: "Exception Occurred : " + error.message
            });
        }
        
    }

    private async initializeComponent() {
        try{
            await this.getMasterDataInformation();
            if (this.props && this.props.invokedItem) {
                try {
                    await this.getInvokedNominationItem(this.props.invokedItem);
                }
                catch (e) {
                    console.log("Error in getting invoked item details.");
                }
            }

            this.SetValidState();
        } catch (error) {
        this.NominationLoggerComponent.Error({
            WebPartName: "PD Nomination",
            ComponentName: "Intake Submission",
            MethodName: "initializeComponent",
            Message: "Exception Occurred : " + error.message
        });
    }
    }
    private async getMasterDataInformation() {
        const masterFormData = await this._handleAsync(this.NominationLibMasters.getMasterDetails());
        this.setState({
            masterListData: masterFormData
        });
    }


    private async getInvokedNominationItem(item: INominationListViewItem) {
        try{
            let itemData: IAllNominationDetails = await this._handleAsync(this.NominationLibComponent.getNominationDetails(item.id, item.nominee, { role: AllRoles.NOMINATOR }));
            const nomineeFormAttach: IAttachment[] = itemData && itemData.nominationAttachments && itemData.nominationAttachments.filter((element) => { return element.attachmentType === "Nomination Form"; });
            const otherAttach: IAttachment[] = itemData && itemData.nominationAttachments && itemData.nominationAttachments.filter((element) => { return element.attachmentType !== "Nomination Form" && element.attachmentType !== null; });

            this.setState(() => ({
                itemDetails: itemData,
                intakeNomination: itemData && itemData.intakeNomination,
                NominationFormAttachment: nomineeFormAttach,
                NominationOtherAttachments: otherAttach,
                isFormStatus: itemData && itemData.intakeNomination && itemData.intakeNomination.nominationStatus,
                isLARequired : itemData.intakeNomination ? itemData.intakeNomination.nomineeDesignation && itemData.intakeNomination.isStatusGrantedAfter2016 ? false : true : false,
                detailsNominationReferences: itemData.nominationReferences,
            }), () => {
                this.SetValidState();
            });
        } catch (error) {
            this.NominationLoggerComponent.Error({
                WebPartName: "PD Nomination",
                ComponentName: "Intake Submission",
                MethodName: "getInvokedNominationItem",
                Message: "Exception Occurred : " + error.message
            });
        }
    }

    private dropDownListObject(items: string[]) {
        return items.sort().map(item => { return { "key": item, "text": item }; });
    }

    private dropDownListSubCategoryObject(items: string[]) {
        const {intakeNomination, existingSubcategory} = this.state; 
        let matchingSubCategories: any;
        if(intakeNomination && intakeNomination.pdDiscipline == "Employee Benefits"  && intakeNomination.financeUserID && existingSubcategory){
             matchingSubCategories = this.state.existingSubcategory.filter(cat => cat["pdDesignation"] == this.state.intakeNomination.pdStatus && cat["_employeeSubCategory"]);
        } 
        return items.sort().map(item => { 
            return  { "key": item, 
                     "text": item, 
                     disabled: matchingSubCategories && matchingSubCategories.length > 0 && matchingSubCategories.filter(matchVal => matchVal["_employeeSubCategory"] === item).length > 0  ? true: false
                    };  
        });
    }

    private async getEmployeeExistingSubCategory() {
        try{    
            const {intakeNomination} = this.state;
            if(intakeNomination && intakeNomination.pdDiscipline == "Employee Benefits"  && intakeNomination.financeUserID){
                const employeeExistingSubcategory = await this.NominationLibMasters.getEmployeeInformation(intakeNomination.financeUserID);
                this.setState({
                existingSubcategory: employeeExistingSubcategory
                });
            }
        } catch (error) {
            this.NominationLoggerComponent.Error({
                WebPartName: "PD Nomination",
                ComponentName: "Intake Submission",
                MethodName: "getEmployeeExistingSubCategory",
                Message: "Exception Occurred : " + error.message
            });
        }
    }

    private clearReferencesValues() {
        const {intakeNomination} = this.state;
        if((intakeNomination && intakeNomination.pdStatus && intakeNomination.trackCandidateNominated && intakeNomination.trackCandidateNominated == INITIAL_CANDIDATE_NOMINATED[0].key && intakeNomination.pdStatus.toUpperCase() == "APPROVED PROFESSIONAL") || (intakeNomination && intakeNomination.pdStatus && intakeNomination.pdStatus.toUpperCase() == "RECOGNIZED PROFESSIONAL"))
        {
            this.setState((prevState) => ({
                detailsNominationReferences: [],
                loading: false
            }), () => {
                this.SetValidState();
            });
        }
    }

    private onNominationFilesChanged = (files: IAttachment[]) => {
        this.setState((prevState) => ({
            NominationFormAttachment: files
        }));   
        this.SetValidState();
    }


    private onOtherFilesChanged = (files: IAttachment[]) => {
        this.setState({
            NominationOtherAttachments: files
        });
    }

    private _checkIfNomineeStatusInProgressBasedOnDesignationAndDiscipline = async() => {
        try{
            const {intakeNomination} = this.state;
            if(intakeNomination && intakeNomination.nominee && intakeNomination.nominee.id && intakeNomination.pdStatus && intakeNomination.pdDiscipline)
            {
                const validNominee: INomineeExist = await this.NominationLibMasters.checkIfValidNomineeWithDiscAndPDStatus(intakeNomination.financeUserID, intakeNomination.pdStatus, intakeNomination.pdDiscipline);

                this.setState((prevState) => ({
                    isValidNominee: validNominee.isNomineeExist,
                    isValidNomineeEPNominator: validNominee.EPNominator
                }), () => {
                    this.SetValidState();
                });
            }
        } catch (error) {
            this.NominationLoggerComponent.Error({
                WebPartName: "PD Nomination",
                ComponentName: "Intake Submission",
                MethodName: "getEmployeeExistingSubCategory",
                Message: "Exception Occurred : " + error.message
            });
        }

        //  this.validateNotificationPhaseRequiredField(event.target.id); 
    }

    public initializeIntakeFormPanel() {
        const { intakeNomination, isFormStatus, masterListData, NominationFormAttachment, itemDetails,isLARequired, isReferencesValid, detailsNominationReferences } = this.state;
        const isDisabled = isFormStatus === NominationStatus.DraftByNominator || isFormStatus === NominationStatus.RequireAdditionalDetails || isFormStatus === undefined ? false : true;
        const IsReq = isReferencesValid && isReferencesValid.isRequired ? "text-label ms-Label ms-Dropdown-label root-306 required":"text-label ms-Label ms-Dropdown-label root-306";
        const epNominationString = this.props.invokedItem
            && intakeNomination
            && intakeNomination.epNominators !== undefined
            && intakeNomination.epNominators.length > 0 ?

            intakeNomination.epNominators.reduce((prevVal, currVal: any, idx) => {
                prevVal.push("/" + currVal.title);
                return prevVal; // *********  Important ******
            }, [])
            : [];

            const referencesString = this.props.invokedItem
            && detailsNominationReferences
            && detailsNominationReferences !== undefined
            && detailsNominationReferences.length > 0 ?

            detailsNominationReferences.reduce((prevVal, currVal: any, idx) => {
                let userDisplayName = currVal.referencesUser ? "/" + currVal.referencesUser.title : "/";
                prevVal.push(userDisplayName);
                return prevVal; // *********  Important ******
            }, [])
            : [];

        const ddlProfessionalDesignation = masterListData && masterListData.professionalDesignation ? this.dropDownListObject(masterListData.professionalDesignation.filter((item: any) => item._professionalDesignationTitle !== "Provisional Signature Authority").map((title:any) =>  title._professionalDesignationTitle)) : [];
        const ddlDiscipline = masterListData && masterListData.discipline ? this.dropDownListObject(masterListData.discipline.filter(disc => disc["_disciplineName"] !== "JOIN").map((friendlyName:any) =>  friendlyName._disciplineFriendlyName)) : [];
        const ddlPDSubCategory = masterListData && masterListData.pdSubCategory ? this.dropDownListSubCategoryObject(masterListData.pdSubCategory.map((category:any) =>  category._pdSub)) : [];
        const ddlLanguage = masterListData && masterListData.language ? this.dropDownListObject(masterListData.language.map((category:any) =>  category._langText)) : [];

        //const isApprovedProfessional_HighestCredentialedProfessional =  intakeNomination && intakeNomination.pdStatus === "Approved Professional" && intakeNomination.trackCandidateNominated === INITIAL_CANDIDATE_NOMINATED[0].key ? false:true;
       

        // const ddlAttachmentType = masterListData && masterListData.attachmentType ? this.dropDownListObject(masterListData.attachmentType) : [];

        const _onChangeForDropDownControls = (event: any, selectedOption: IDropdownOption,) => {
            const ddlArray: Array<string> = intakeNomination && intakeNomination[event.target.id] ? intakeNomination[event.target.id] : [];
            if (selectedOption) {
                this.setState((prevState) => ({
                    intakeNomination: {
                        ...prevState.intakeNomination,
                        [event.target.id]: selectedOption.selected ? [...ddlArray, selectedOption.key as string] : ddlArray.filter(key => key !== selectedOption.key)
                    }
                }), () => {
                    this.SetValidState();
                });
            }
        };

        const _onChangeForDropDownStringTypeControls = (event, selectedOption: IDropdownOption,) => {
            //const ddlSelectedItem: string = selectedOption.text;
           
            this.setState((prevState) => ({
                intakeNomination: {
                    ...prevState.intakeNomination,
                    [event.target.id]: selectedOption.key as string
                },
            }), () => {
                this._checkIfNomineeStatusInProgressBasedOnDesignationAndDiscipline();
                this.getEmployeeExistingSubCategory();
                this.SetValidState();
            });
            if(event.target.id == "trackCandidateNominated")
            {
                this.clearReferencesValues();
            }
           
            //  this.validateNotificationPhaseRequiredField(event.target.id); 
        };

    
        const _onChangeIntakeNotes = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newText: string): void => {
            const notesContent: string = newText;
            this.setState((prevState) => ({
                intakeNomination: {
                    ...prevState.intakeNomination,
                    intakeNotes: notesContent
                }
            }), () => {
                this.SetValidState();
            });

        };

        const _onChangeBillingCode = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newText: string): void => {
            const billingCodeContent: string = newText;
            this.setState((prevState) => ({
                intakeNomination: {
                    ...prevState.intakeNomination,
                    billingCode: billingCodeContent
                }
            }), () => {
                this.SetValidState();
            });

        };
        
        const _onChangeRPCertification = (event, checked?: boolean) => {
            const isChecked: boolean = checked;
            this.setState((prevState) => ({
                intakeNomination: {
                    ...prevState.intakeNomination,
                    rpCertification: isChecked
                }
            }), () => {
                this.SetValidState();
            });
        };

        return (
            <React.Fragment>
                {this.state.loading ? <SpinnerComponent text={"Loading..."} /> : ""}
                <div className="ms-Panel-scrollableContent scrollableContent-561" data-is-scrollable="true">
                    <div className="ms-Panel-content content-562">
                        <div className={[styles.row, styles.rowEnd].join(" , ")}>
                            <div className={styles.column6}><span className={styles.header}>Submission Form</span> </div>
                        </div>
                        <div className={styles.row}>
                            <div className={styles.column12}>Complete this form to start the Professional Designation nomination process for a candidate. Visit the <a href="https://milliman.sharepoint.com/sites/ProfDesignationNominationSupport/SitePages/ProfDesNominationForm.aspx" data-interception="off" target="_blank" rel="noopener noreferrer">support site </a> for more information.</div>
                        </div>
                        <div className={[styles.row, styles.rowEnd].join(" , ")}>
                            <div className={styles.column12}><span className={styles.mandatInfo}><strong>Fields marked (<span className={styles.star}>*</span>) are mandatory</strong></span> </div>
                        </div>
                        <div className={styles.row}>
                            <div className={styles.column12 + ' text-label'}><span className={styles.heading}>Candidate</span></div>
                        </div>
                        <div className={styles.row}>
                            <div className={styles.column3 + ' text-label'}>
                                {intakeNomination && intakeNomination.nominee && intakeNomination.id && intakeNomination.id != 0 ? <PeoplePicker
                                    context={this.props.context}
                                    titleText="Nominee Name"
                                    personSelectionLimit={1}
                                    showtooltip={true}
                                    required={true}
                                    disabled={true}
                                    showHiddenInUI={false}
                                    defaultSelectedUsers={intakeNomination && intakeNomination.nominee && intakeNomination.nominee.title ? ["/" + intakeNomination.nominee.title.toString()] : null}
                                    principalTypes={[PrincipalType.User]}
                                    resolveDelay={1000}
                                    ensureUser={true}
                                /> : <PeoplePicker
                                    context={this.props.context}
                                    titleText="Nominee Name"
                                    personSelectionLimit={1}
                                    showtooltip={true}
                                    required={true}
                                    disabled={isDisabled}
                                    onChange={this._getNomineeInformation}
                                    errorMessage={this.state && intakeNomination && intakeNomination.nominee && !this.state.isValidNominee ? this.state.isValidNomineeEPNominator + " already started the nomination process for this employee. Please reach out to them for more details. " : null}
                                    showHiddenInUI={false}
                                    defaultSelectedUsers={intakeNomination && intakeNomination.nominee && intakeNomination.nominee.title ? ["/" + intakeNomination.nominee.title.toString()] : null}
                                    principalTypes={[PrincipalType.User]}
                                    resolveDelay={1000}
                                    ensureUser={true}
                                />}
                                {/* {this.state.loading &&
                                <Spinner label='Loading...' ariaLive='assertive' /> }*/}

                            </div>
                            <div className={styles.column3 + ' text-label'}>
                                <TextField label="Practice"
                                    disabled={true}
                                    value={intakeNomination && intakeNomination.nomineePractice ? intakeNomination.nomineePractice : ""}
                                />
                                {/* {this.state.loading &&
                                <Spinner label='Loading...' ariaLive='assertive' />
                            } */}
                            </div>
                            <div className={styles.column3 + ' text-label'}>
                                <TextField label="Office"
                                    disabled={true}
                                    value={intakeNomination && intakeNomination.nomineeOffice ? intakeNomination.nomineeOffice : ""}
                                />
                            </div>
                            <div className={styles.column3 + ' text-label'}>
                                <TextField label="Discipline"
                                    disabled={true}
                                    value={intakeNomination && intakeNomination.nomineeDiscipline ? intakeNomination.nomineeDiscipline : ""}
                                />
                            </div>
                        </div>
                        <div className={[styles.row, styles.rowEnd].join(" , ")}>
                          {/*   <div className={styles.column3 + ' text-label'}>
                                <TextField label=" Professional Designation"
                                    disabled={true}
                                    hidden={true}
                                    value={intakeNomination && intakeNomination.nomineeDesignation ? intakeNomination.nomineeDesignation : ""}
                                />
                            </div> */}

                           {/*  <div className={styles.column3 + ' text-label'}>
                                <TextField label="Is Status After 2016"
                                    disabled={true}
                                    value={intakeNomination && intakeNomination.isStatusGrantedAfter2016 ? "Yes" : "No"}
                                    hidden={true}
                                />
                            </div> */}
                        </div>

                        <div className={styles.row}>
                            <div className={styles.column12 + ' text-label'}><span className={styles.heading}>Nominate For</span></div>
                        </div>
                        <div className={styles.row}>
                            <div className={styles.column3 + ' text-label'}>
                                <Dropdown placeholder="Select Professional Designation"
                                    disabled={isDisabled}
                                    id={"pdStatus"}
                                    label="Professional Designation"
                                    options={ddlProfessionalDesignation}
                                    //options={this.masterDetails.professionalDesignation.length > 0 ? this.masterDetails.professionalDesignation.map(stringText => ({key: stringText.code, text:stringText.title})): []} 
                                    required={true}
                                    onChange={_onChangeForDropDownStringTypeControls}
                                    defaultSelectedKey={intakeNomination && intakeNomination.pdStatus ? intakeNomination.pdStatus : ""}
                                />
                            </div>
                            <div className={styles.column3 + ' text-label'}>
                                <Dropdown placeholder="Select PD Discipline"
                                    disabled={isDisabled}
                                    label="PD Discipline"
                                    id={"pdDiscipline"}
                                    options={ddlDiscipline}
                                    //options={this.masterDetails.discipline.length > 0 ? this.masterDetails.discipline.map(stringText => ({key: stringText.abbreviation, text:stringText.friendlyName})): []} 
                                    required={true}
                                    onChange={
                                        _onChangeForDropDownStringTypeControls
                                    }
                                    selectedKey={intakeNomination && intakeNomination.pdDiscipline ? intakeNomination.pdDiscipline : ""}
                                />
                            </div>
                            {intakeNomination && intakeNomination.pdDiscipline === "Employee Benefits" ?
                                <div className={styles.column3 + ' text-label'}>
                                    <Dropdown placeholder="Select Subcategory"
                                        disabled={isDisabled}
                                        label="Subcategory"
                                        id={"pdSubcategory"}
                                        options={ddlPDSubCategory}
                                        multiSelect
                                        required={true}
                                        onChange={_onChangeForDropDownControls}
                                        //defaultSelectedKeys={['Pension - Corporate', 'Pension - Public Sector']}

                                        defaultSelectedKeys={intakeNomination && intakeNomination.pdSubcategory ? intakeNomination.pdSubcategory : []}
                                    />
                                </div>
                            : ""}
                            <div className={styles.column3 + ' text-label'}>
                                <PeoplePicker
                                    context={this.props.context}
                                    titleText="EP Nominator(s)"
                                    key={"epNominators"}
                                    personSelectionLimit={5}
                                    showtooltip={true}
                                    required={true}
                                    disabled={isDisabled}
                                    onChange={this._getEPNominatorInformation}
                                    ensureUser={true}
                                    showHiddenInUI={false}
                                    principalTypes={[PrincipalType.User]}
                                    //this.state.intakeNomination.epNominators.map(String).join("/").toString()
                                    defaultSelectedUsers={intakeNomination && intakeNomination.epNominators && intakeNomination.epNominators.length > 0 ? epNominationString : []}
                                    resolveDelay={1000} />
                            </div>
                        </div>
                        <div className={[styles.row, styles.rowEnd].join(" , ")}>
                            <div className={styles.column6   + ' text-label'}>
                                <Dropdown placeholder="Select Language"
                                    disabled={isDisabled}
                                    id={"proficientLanguage"}
                                    multiSelect
                                    label="List any language in which the candidate is proficient to perform work"
                                    options={ddlLanguage}
                                    required={true}
                                    onChange={_onChangeForDropDownControls}
                                    defaultSelectedKeys={intakeNomination && intakeNomination.proficientLanguage ? intakeNomination.proficientLanguage : []}

                                />
                            </div>
                            
                            <div className={styles.column4 + ' text-label'}>
                                <Dropdown placeholder="Select track is the candidate nominated for"
                                    disabled={isDisabled}
                                    label="Under which track is the candidate nominated ?"
                                    id={"trackCandidateNominated"}
                                    options={INITIAL_CANDIDATE_NOMINATED}
                                    //options={this.masterDetails.discipline.length > 0 ? this.masterDetails.discipline.map(stringText => ({key: stringText.abbreviation, text:stringText.friendlyName})): []} 
                                    required={true}
                                    onChange={
                                        _onChangeForDropDownStringTypeControls
                                    }
                                    selectedKey={intakeNomination && intakeNomination.trackCandidateNominated ? intakeNomination.trackCandidateNominated : ""}
                                />
                            </div>
                            {(intakeNomination && intakeNomination.pdStatus && intakeNomination.trackCandidateNominated && intakeNomination.pdStatus.toUpperCase() === "SIGNATURE AUTHORITY" && intakeNomination.trackCandidateNominated === INITIAL_CANDIDATE_NOMINATED[2].key)?
                                <div className={styles.column2  + ' text-label'}>
                                    <TextField label="Billing code"
                                        required={true}
                                        disabled={false}
                                        onChange={_onChangeBillingCode} 
                                        placeholder='Enter the billing code'
                                        value={intakeNomination && intakeNomination.billingCode ? intakeNomination.billingCode : ""}
                                    />
                                </div>
                                :""
                            }
                          
                            {
                            /*
                            <div className={styles.column3 + ' text-label'}>
                                <Toggle
                                    checked={intakeNomination && intakeNomination.isProductPerson}
                                    label={
                                        <div>
                                            Is the candidate nominated under PTPAC Professionals Nominations guidelines?
                                        </div>
                                    }
                                    disabled={isDisabled}
                                    inlineLabel
                                    id={"isProductPerson"}
                                    onText="Yes"
                                    offText="No"
                                    onChange={_onChangeIsProductPerson}
                                />
                            </div>
                            */
                            }
                        </div>
                        {isReferencesValid && isReferencesValid.isHide === false ?
                        <>
                        <div className={styles.row}>
                            <div className={styles.column12 + ' text-label'}><span className={styles.heading}>References</span></div>
                        </div>
                        <div className={[styles.row, styles.rowEnd].join(" , ")}>
                            <div className={styles.column12 + ' text-label'}>
                                <label className={IsReq} id="Dropdown62-label">
                                    A minimum of two references is required for an AP nomination and a minimum of three references is required for a SA/QR nomination. References are not required for a first-time LSA nomination and are encouraged with a renewal LSA nomination.
                                    {
                                        isReferencesValid && isReferencesValid.isRequired?<span className={styles.star}>*</span>:""
                                    }  
                                </label>
                            </div>
                            <div className={styles.column6 + ' text-label'}>
                                <PeoplePicker
                                    context={this.props.context}
                                    personSelectionLimit={5}
                                    showtooltip={true}
                                    required={isReferencesValid && isReferencesValid.isRequired?true:false}
                                    key={"referee"}
                                    disabled={isDisabled}
                                    onChange={this._getRefereeInformation}
                                    ensureUser={true}
                                    showHiddenInUI={false}
                                    principalTypes={[PrincipalType.User]}
                                    defaultSelectedUsers={detailsNominationReferences && detailsNominationReferences.length > 0 ? referencesString : []}
                                    resolveDelay={1000} />
                            </div>
                        </div>
                        </>
                        : ""}
                        {isLARequired &&
                            <div>
                                <div className={styles.row}>
                                    <div className={styles.column12 + ' text-label'}><span className={styles.heading}>Local Reviewer</span></div>
                                </div>
                                <div className={[styles.row, styles.rowEnd].join(" , ")}>
                                    <div className={styles.column12 + ' text-label'}>
                                        <label className="text-label ms-Label ms-Dropdown-label root-306" id="Dropdown62-label">
                                            Who will be responsible for reviewing the employee agreements <a data-interception='off' target='_blank' rel='noopener noreferrer'>and for updating the employee number? *</a>
                                        </label>
                                    </div>
                                    <div className={styles.column6 + ' text-label'}>
                                        <PeoplePicker
                                            context={this.props.context}
                                            personSelectionLimit={1}
                                            showtooltip={true}
                                            required={true}
                                            key={"Assignee"}
                                            disabled={isDisabled}
                                            onChange={this._getNominatorInformation}
                                            ensureUser={true}
                                            showHiddenInUI={false}
                                            principalTypes={[PrincipalType.User]}
                                            defaultSelectedUsers={itemDetails && itemDetails.nominationDetailsByLA && itemDetails.nominationDetailsByLA.assignee ? ["/" + itemDetails.nominationDetailsByLA.assignee.title.toString()] : []}
                                            resolveDelay={1000} />
                                    </div>
                                </div>
                            </div>
                        }
                        {intakeNomination && intakeNomination.pdStatus && intakeNomination.pdStatus !== PdStatus.RP ?
                        <>
                            <div className={styles.row}>
                                <div className={styles.column6 + ' text-label'}><span className={styles.heading}>Comments</span></div>
                            </div>
                            <div className={[styles.row, styles.rowEnd].join(" , ")}>
                                <div className={styles.column12 + ' text-label'}>
                                    <TextField label="" multiline rows={6} disabled={isDisabled}
                                        onChange={_onChangeIntakeNotes} placeholder='Enter comments to your discipline Quality Coordinator '
                                        id={"intakeNotes"}
                                        value={intakeNomination && intakeNomination.intakeNotes ? intakeNomination.intakeNotes : ""}
                                    />
                                </div>
                            </div>
                        </>
                        :""
                        }
                        {intakeNomination && intakeNomination.pdStatus && intakeNomination.pdStatus !== PdStatus.RP ?
                            <div>
                                <div className={styles.row}>
                                    <div className={styles.column6 + ' text-label'}><span className={styles.heading}>Nomination Form <span className={styles.star}>*</span></span></div>
                                </div>
                                <div className={styles.column12 + ' text-label'}>
                                    <FileUploader
                                        onFilesChanged={(fileItem) => { this.onNominationFilesChanged(fileItem); }}
                                        docType={"Nomination Form"}
                                        context={this.props.context}
                                        disabled={isDisabled}
                                        role={AllRoles.NOMINATOR}
                                        onDocumentDelete={(fileItem) => { this.delNominationForm(fileItem); }}
                                        files={NominationFormAttachment && NominationFormAttachment.length > 0 && NominationFormAttachment.map((attachment: IAttachment) => {
                                            return attachment;
                                        })}
                                    >
                                    </FileUploader>
                                </div>
                                <div className={styles.row}>
                                    <div className={styles.column6 + ' text-label'}><span className={styles.heading}>Other Attachments</span></div>
                                </div>
                                <div className={styles.row}>
                                    <div className={styles.column12 + ' text-label'}>
                                        <FileUploader
                                            onFilesChanged={(fileItem) => { this.onOtherFilesChanged(fileItem); }}
                                            context={this.props.context}
                                            docType={this.state.attachmentType && this.state.attachmentType}
                                            disabled={isDisabled}
                                            role={AllRoles.NOMINATOR}
                                            onDocumentDelete={!isDisabled ? (fileItem) => { this.delOtherAttachments(fileItem);} : null}
                                            files={this.state.NominationOtherAttachments.length > 0 && this.state.NominationOtherAttachments.map((attachment: IAttachment) => {
                                                return attachment;
                                            })}
                                        >
                                        </FileUploader>
                                    </div>
                                </div>
                            </div>
                            : ""}
                        {intakeNomination && intakeNomination.pdStatus && intakeNomination.pdStatus === PdStatus.RP ?
                            <div className={styles.row}>
                                <div className={styles.column12 + ' text-label'}>
                                    <Toggle
                                        checked={intakeNomination && intakeNomination.rpCertification}
                                        label={
                                            <div>
                                                I affirm that the nominee meets the criteria for <a
                                                    href="https://milliman.sharepoint.com/sites/GCSLegal/New Policies Library/Recognized Professional Status.pdf?action=default&mobileredirect=true&DefaultItemOpen=1"
                                                    data-interception="off" target="_blank" rel="noopener noreferrer">Recognized Professional status</a>
                                            </div>
                                        }
                                        id={"rpCertification"}
                                        disabled={isDisabled}
                                        inlineLabel
                                        onText="Yes"
                                        offText="No"
                                        onChange={
                                            _onChangeRPCertification
                                            //this.validateNotificationPhaseRequiredField(FormField.isProductPerson);
                                        }
                                    />
                                </div>
                            </div>
                            : ""}
                    </div>
                </div></React.Fragment>);
    }

    private _handleAsync = (promise): Promise<any> => {
        return promise
            .then(data => (data))
            .catch(error => Promise.resolve(error));
    }


    private _clickPanel(e) {
        e.preventDefault();
    }


    public footerRender() {
        const { isFormStatus, isSaveValid, isSubmitValid, intakeNomination } = this.state;
        const isDisabled = isSaveValid && (isFormStatus === NominationStatus.DraftByNominator || isFormStatus === NominationStatus.RequireAdditionalDetails || !isFormStatus) && intakeNomination && intakeNomination.nominee !== null ? false : true;
        const isDeleteVisible = isFormStatus === NominationStatus.DraftByNominator ? true : false;
        const isSubmitDisabled = isSubmitValid && (isFormStatus === NominationStatus.DraftByNominator || isFormStatus === NominationStatus.RequireAdditionalDetails || !isFormStatus) && intakeNomination && intakeNomination.nominee !== null ? false : true;

        return (
            <div className={styles.footerRow}>
                <div className={styles.column12}>
                    <Stack horizontal tokens={stackTokens}>
                        <DefaultButton disabled={isDisabled} onClick={() => { this.processIntakeForm(0); }} text={GENERICBUTTONSACTIONS.SAVE} />
                        <DefaultButton disabled={isSubmitDisabled} onClick={() => { this.processIntakeForm(1); }} text= {GENERICBUTTONSACTIONS.SUBMIT} />
                        {isDeleteVisible && <DefaultButton onClick={() => { this._setIsDialogVisible(true); }} text={GENERICBUTTONSACTIONS.DELETE} />}
                        <DefaultButton onClick={this.close && this.props.onDismiss} disabled={false} text="Close" />
                    </Stack>
                </div>
            </div>);
    }




    public render(): React.ReactElement<IIntakePanelProps> {

        return (
            // <div className={styles.registrationInfoText + " " + styles.row} >
            //     {//this.props.id == 0 ? <div className={styles.column12}><a href="#" onClick={(e) => { this.openPanel(e); }}>+ Start the registration process for your product, packaged solution, or expansion</a></div> : <div className={styles.column12}><PrimaryButton text="Edit" type="button" onClick={(e) => { this.openPanel(e); }} /></div
            //     }
            this.state && this.state.isOpen && < Panel
                isOpen={this.state.isOpen}
                onDismiss={this.close && this.props.onDismiss}
                onDoubleClick={(e) => this._clickPanel(e)}
                type={PanelType.extraLarge}
                closeButtonAriaLabel="Close"
                isFooterAtBottom={true}
                onRenderFooterContent={() => this.footerRender()}>
                <div className={styles.registrationForm + ' ' + 'formDisplay'}>
                    <div className={styles.container}>
                        {this.state.isOpen ? this.initializeIntakeFormPanel() : ""}
                        {this._getMessage()}    
                    </div>
                </div>

                <Dialog
                    hidden={!this.state.isConfirmationDialogVisible}
                    dialogContentProps={dialogContentProps}
                    modalProps={dialogModalProps}
                >
                    <DialogFooter>
                        <PrimaryButton onClick={() => { this._setIsDialogVisible(false); this.processIntakeForm(2); }} text="Yes" />
                        <DefaultButton onClick={() => { this._setIsDialogVisible(false); }} text="No" />
                    </DialogFooter>
                </Dialog>
            </Panel >
            // </div >
        );
    }

    @autobind
    private _setIsDialogVisible(value) {
        this.setState({
            isConfirmationDialogVisible: value
        });
    }
    private comparer(otherArray: IAttachment[], type?: string) {
        return current => {
            return otherArray.filter(other => {
                return other.attachmentType == type;
            }).length > 0;
        };
    }

    @autobind
    private close() {
        //this._onCloseTimer = setTimeout(this._onClose.bind(this), parseFloat(styles.duration));
        this.setState({
            isOpen: false,
            actionStatus: STATUS.DEFAULT,
            isMessageDialogVisible: false
        });
        this.props.onDismiss();
    }
    

    private delNominationForm = async (file: string) => {
        const {itemDetails, NominationFormAttachment} = this.state;
        let removeFromNominationFormAttachments = null;
        if(this.state.itemDetails && this.state.itemDetails.nominationAttachments.length > 0){
            const subfolderName: string = this.state.itemDetails.nominationAttachments.filter(a => a.attachmentName == file)[0].attachmentBy;
            removeFromNominationFormAttachments = this.state.itemDetails.nominationAttachments.reduce((p,c) => (c.attachmentName !== file && p.push(c),p),[]);
            await this.NominationLibComponent.deleteFile(this.state.itemDetails, { role: AllRoles.NOMINATOR },subfolderName, file);
            itemDetails.nominationAttachments = removeFromNominationFormAttachments;
           
        }
        else if(NominationFormAttachment && NominationFormAttachment.length > 0){
            removeFromNominationFormAttachments = NominationFormAttachment.filter(a => a.attachmentName != file);
        }
        this.setState(() => ({
            NominationFormAttachment: removeFromNominationFormAttachments,
        }), () => {
            this.SetValidState();
        });
    }

    private delOtherAttachments = async (file: string) => {
        const {itemDetails, NominationFormAttachment} = this.state;
        let removeOtherAttachments = null;
        if(this.state.itemDetails && this.state.itemDetails.nominationAttachments.length > 0){
            const subfolderName: string = this.state.itemDetails.nominationAttachments.filter(a => a.attachmentName == file)[0].attachmentBy;
            removeOtherAttachments = this.state.itemDetails.nominationAttachments.reduce((p,c) => (c.attachmentName !== file && p.push(c),p),[]);
            await this.NominationLibComponent.deleteFile(this.state.itemDetails, { role: AllRoles.NOMINATOR },subfolderName, file);
            itemDetails.nominationAttachments = removeOtherAttachments;
           
        }
        else if(NominationFormAttachment && NominationFormAttachment.length > 0){
            removeOtherAttachments = NominationFormAttachment.filter(a => a.attachmentName != file);
        }
        this.setState(() => ({
            NominationOtherAttachments: removeOtherAttachments,
        }), () => {
            this.SetValidState();
        });
    }

    private processAttachments(updatedAttachments: IAttachment[], oldAttachments: IAttachment[]): IAttachment[] {
        if (oldAttachments) {
            oldAttachments.forEach((f, i) => {
                let inUpdatedAttachments = updatedAttachments ? updatedAttachments.filter((m) => { if (m.attachmentName == f.attachmentName) return true; }) : [];
                if (!inUpdatedAttachments || inUpdatedAttachments.length == 0) {
                    updatedAttachments.push({
                        attachmentBy: f.attachmentBy,
                        attachmentName: f.attachmentName,
                        attachmentType: f.attachmentType,
                        id: 0
                    });
                }

            });
        }
        return updatedAttachments;
    }

    private async processIntakeForm(action: number) {
        let { intakeNomination, itemDetails, detailsNominationReferences } = this.state;
        let qcUsers = null;
        const mergeConcat = (...arrays) => [].concat(...arrays.filter(Array.isArray));
        const notNullReferencesCollection = detailsNominationReferences.filter(refer => refer.id != 0 || refer.referencesUser !== null);

        if(intakeNomination.pdDiscipline && action === 1)
        {
            qcUsers = await this._handleAsync(this.NominationListLibComponent.getQCDisciplineUsers(intakeNomination.pdDiscipline));
        }  
        this.setState({
            loading: true,
            nominationReviewersUsers: qcUsers
        });

        intakeNomination = {
            ...intakeNomination,
            financeUserID: intakeNomination.financeUserID.toString(),
            
        };
       
        if (action === 0) {
            intakeNomination = {
                ...intakeNomination,
                draftDate: CommonMethods.getSPFormatDate(new Date()).toString(),
                nominationStatus: NominationStatus.DraftByNominator,
            };
        }
        else if (action === 1 && intakeNomination.submissionDate) {
            intakeNomination = {
                ...intakeNomination,
                reSubmissionDate: CommonMethods.getSPFormatDate(new Date()).toString(),
                nominationStatus: NominationStatus.SubmittedByNominator
            };
        }
        else if (action == 1) {
            intakeNomination = {
                ...intakeNomination,
                submissionDate:CommonMethods.getSPFormatDate(new Date()),
                nominationStatus: NominationStatus.SubmittedByNominator
            };
        }
        else if (action == 2) { //if deleted
            intakeNomination = {
                ...intakeNomination,
                nominationStatus: NominationStatus.Deleted
            };
        }
        
        let updatedAttachments: IAttachment[] = [];
        if (action != 2) {
            updatedAttachments = mergeConcat(this.state.NominationFormAttachment, this.state.NominationOtherAttachments);
            let oldAttachments = itemDetails && itemDetails.nominationAttachments;
            updatedAttachments = this.processAttachments(updatedAttachments, oldAttachments);
        }
        let allNominationDetails: IAllNominationDetails = {
            ...itemDetails,
            intakeNomination: intakeNomination,
            nominationAttachments: updatedAttachments,
            nominationReferences: notNullReferencesCollection
        };

        await this.processData(allNominationDetails, action);
    }

    private _getNomineeProfessionalDesignation = async () => {
        const {intakeNomination} = this.state;
        const nomineeFinanceUserID = parseInt(intakeNomination.financeUserID);
        let nomineePDAPIData: IProfessionalDesignationDetailed[] = null;
        let clearValue = false;
        if (nomineeFinanceUserID > 0) {
            nomineePDAPIData = await this._handleAsync(this.NominationListLibComponent.getProfessionalDesignationsByFinanceUserId(nomineeFinanceUserID));
        }
        else {
            clearValue = true;
        }
        // this.setState({ loading: true }, async () => {


        if (nomineePDAPIData) {
           const result: IProfessionalDesignationDetailed[] = nomineePDAPIData.map((elem: IProfessionalDesignationDetailed) => (
                {    
                    id: elem.id,
                    financeUserId: elem.financeUserId,
                    designationId:elem.financeUserId,
                    pdSubategoryId:elem.pdSubategoryId,
                    discipline:elem.discipline,
                    grantedOn:elem.grantedOn,
                    removedOn:elem.removedOn,
                    level:elem.level,
                    restrictions:elem.restrictions,
                    restrictionDate:elem.restrictionDate,
                    professionalDesignation:elem.professionalDesignation,
                    subCategory:elem.subCategory,
                    code:elem.code,
                    abbreviation:elem.abbreviation,
                    friendlyName:elem.friendlyName,
                    created:elem.created,
                    modified:elem.modified,
                    createdBy:elem.createdBy,
                    modifiedBy:elem.modifiedBy,
                } 

                ));
                this.setState((prevState) => ({
                    pdNominationDetailed: result
                }));
        }
            
           
    }

    private async _UpdateEmployeeProfessionalDesignation(): Promise<boolean> {
        try {
            const {masterListData, intakeNomination, pdNominationDetailed, grantedOn} = this.state;
            const profDestId =  intakeNomination.pdStatus ? masterListData.professionalDesignation.filter((_pd: any) => _pd._professionalDesignationTitle == intakeNomination.pdStatus)[0]["_professionalDesignationId"] : null;
            const subcategoryId =  intakeNomination.pdSubcategory ? masterListData.pdSubCategory.filter((_sc: any) => intakeNomination.pdSubcategory.indexOf(_sc._pdSub) > -1) : null;
                                                                  
            const disciplineId =  intakeNomination.pdDiscipline ? masterListData.discipline.filter((_disc: any) => _disc._disciplineFriendlyName == intakeNomination.pdDiscipline)[0]["_disciplineId"] : null;
            const proficientLanguagesIds =  intakeNomination.proficientLanguage ? masterListData.language.filter((_lang: any) => intakeNomination.proficientLanguage.indexOf(_lang._langText) > -1) : null;
            

            const nomineeFinanceUserID = parseInt(intakeNomination.financeUserID);
            const isNomineePDStatusNew = pdNominationDetailed && pdNominationDetailed.length > 0 ? pdNominationDetailed.filter((_disc: any) => _disc.friendlyName == intakeNomination.pdDiscipline && _disc.professionalDesignation == intakeNomination.pdStatus) : []; 
            
            const arrProficientLanguages =  proficientLanguagesIds && proficientLanguagesIds.length > 0 ? proficientLanguagesIds.map(lang => ({ id: null, financeUserId: nomineeFinanceUserID, proficientLaguageId: lang["_langId"], isDelete: false})) : [{id: null, financeUserId: 0, proficientLaguageId: null, isDelete: false }];
            const arrProfessionalDesignation =  profDestId && subcategoryId && subcategoryId.length > 0 ? subcategoryId.map(sub => ({ 
                id: null, 
                financeUserId: nomineeFinanceUserID, 
                designationId: profDestId, 
                pdSubategoryId: sub["_pdId"],
                disciplineId: disciplineId,
                grantedOn: grantedOn,
                removedOn: null,
                level: null,
                isDelete: false})) 
                : 
                [{ 
                    id: null, 
                    financeUserId: nomineeFinanceUserID, 
                    designationId: profDestId, 
                    pdSubategoryId: subcategoryId,
                    disciplineId: disciplineId,
                    grantedOn: grantedOn,
                    removedOn: null,
                    level: null,
                    isDelete: false
                }];
            
            const insertEmployeeObject: IEmployeeUpdateProperties = {
                financeUserId: nomineeFinanceUserID,
                committeeAssignments: [{ id: null, financeUserId: 0, committeeId: null, isDelete: false }],
                proficientLanguages: arrProficientLanguages,
                professionalDesignations: arrProfessionalDesignation,
                shareholders: [{
                    id: null,
                    financeUserId: 0,
                    shareholderLegalId: null,
                    shareholderId: null,
                    date: new Date(),
                    isDelete: false,
                }]
                
            };
            if (nomineeFinanceUserID && isNomineePDStatusNew  && isNomineePDStatusNew.length === 0 && insertEmployeeObject) {
                await this._handleAsync(this.NominationListLibComponent.updateNomineeEmployeeDetails(insertEmployeeObject));
                return Promise.resolve(true);
            }
            return  Promise.resolve(false);
        }   catch (err) {
            if (err instanceof Error) {
                console.error(`Things exploded (${err.message})`);
            }
        }
          
    }
    
    private async processData(intakeNominationData: IAllNominationDetails, action: number) {
        const{nominationReviewersUsers}=this.state;
        const postURL: string = this.Constants.PowerAutomateFlowUrl;


        if (intakeNominationData.intakeNomination) {
            const isPDStatusGrantedBefore2016 = nominationReviewersUsers != null && intakeNominationData.intakeNomination.nomineeDesignation!=null && intakeNominationData.intakeNomination.pdStatus!=PdStatus.RP && this.state.isLARequired == true;
            const isPDStatusGrantedAfter2016 = nominationReviewersUsers != null && intakeNominationData.intakeNomination.nomineeDesignation!=null && intakeNominationData.intakeNomination.pdStatus!=PdStatus.RP && this.state.isLARequired == false;
            
          
            
            const saved: any = await this.NominationLibComponent.saveNominationDetails(intakeNominationData, { role: AllRoles.NOMINATOR }, null,null);

            
            const permissionParameters = CommonMethods.setPermissionOnAttachment(this.props.context,
                saved.attachmentDocumentSetName,
                AllRoles.NOMINATOR,
                this.Constants.SP_LIST_NAMES.NominationDocumentLibraryName,
                this.props.context.pageContext.user.email
            ); 
            

            if(permissionParameters){
                this.EmailNotification.nominationAttachmentPermission(permissionParameters, this.Constants.PermissionPowerAutomateFlowUrl);
                const isSendAnEmail : boolean = NominationStatus.PendingWithQC && (intakeNominationData.intakeNomination.pdStatus && intakeNominationData.intakeNomination.pdStatus.indexOf(PdStatus.RP) == -1);
                const currentWebUrl = this.props.context.pageContext.web.absoluteUrl;
            
                let emailContentForNominee: any = [];
                let actionStatus = STATUS.DEFAULT;
                if (action == 0) {
                    actionStatus = saved ? STATUS.SAVE_SUCCESS : STATUS.SAVE_ERROR;
                }
                else if (action == 1 && saved) {

                    let currentUserEmail = this.props.context.pageContext.user.email;
                    actionStatus = saved ? STATUS.SUBMIT_SUCCESS : STATUS.SUBMIT_ERROR;

                    emailContentForNominee = await this.EmailNotification.getNotificationList(NominationStatus.SubmittedByNominator+'-'+'Nominee', { role: AllRoles.NOMINATOR }, intakeNominationData);
                    
                    if (typeof emailContentForNominee === 'object' && emailContentForNominee !== null && isSendAnEmail && emailContentForNominee[0].IsEnabled) {
                        const body= this.makeEmailBody(intakeNominationData.intakeNomination.nominee.email,emailContentForNominee[0].emailSub,emailContentForNominee[0].emailBody,emailContentForNominee[0].emailCC,AllRoles.NOMINATOR,'',currentWebUrl,[],currentUserEmail);
                        this.EmailNotification.nominationEmail(body, postURL);
                    }
                    //11/3/2021:- Here We are sending Emails to EP Nominators  different Subject and Body//
                    const epNominatorsEmail = intakeNominationData.intakeNomination.epNominators.map((element) => { return element.email; }).join(';');
                    if (intakeNominationData.intakeNomination.epNominators.length > 0) {
                        const emailContentEP = await this.EmailNotification.getNotificationList(NominationStatus.SubmittedByNominator + '-' + AllRoles.EP_NOMINATOR, { role: AllRoles.EP_NOMINATOR }, intakeNominationData);
                        if(emailContentEP && emailContentEP.length > 0 && emailContentEP[0].IsEnabled)
                        {
                            const body= this.makeEmailBody(epNominatorsEmail,emailContentEP[0].emailSub,emailContentEP[0].emailBody,emailContentEP[0].emailCC,AllRoles.EP_NOMINATOR,AllRoles.EP_NOMINATOR,currentWebUrl,[],currentUserEmail);
                            this.EmailNotification.nominationEmail(body, postURL);
                        }
                    }
                    //11/3/2021: Here We have to Send an Email to Local Reviewer of the form //
                
                // if (this.state.isLARequired == true) {
                    if (this.state.isLARequired == true) {
                        const emailContentLA = await this.EmailNotification.getNotificationList(NominationStatus.SubmittedByNominator + '-' + AllRoles.LA, { role: AllRoles.LA }, intakeNominationData);
                        if(emailContentLA && emailContentLA.length > 0 && emailContentLA[0].IsEnabled)
                        {
                            const body= this.makeEmailBody(intakeNominationData.nominationDetailsByLA.assignee.email,emailContentLA[0].emailSub,emailContentLA[0].emailBody,emailContentLA[0].emailCC,AllRoles.LA,AllRoles.LA,currentWebUrl,[],currentUserEmail);
                            this.EmailNotification.nominationEmail(body, postURL);
                        }
                    }

                    //Update Employee Database if Granted status after 2016 and Newly PD Status is RP and ere We have to send Congratulation Message to Selected RP Status

                    if((intakeNominationData.intakeNomination.nomineeDesignation) && (intakeNominationData.intakeNomination.pdStatus==PdStatus.RP) && (intakeNominationData.intakeNomination.isStatusGrantedAfter2016))
                    {
                        const updateDB  =  await this._handleAsync(this._UpdateEmployeeProfessionalDesignation());
                        if(updateDB){
                            const emailContentUpdateDB = await this.EmailNotification.getNotificationList(NominationStatus.PendingWithLocalAdmin+'-'+PdStatus.RP, {role: AllRoles.LA}, intakeNominationData,intakeNominationData.intakeNomination.pdDiscipline);
                            if (typeof emailContentUpdateDB === 'object' && emailContentUpdateDB !== null && emailContentUpdateDB.length > 0 && emailContentUpdateDB[0].IsEnabled) {
                                const body= this.makeEmailBody(intakeNominationData.intakeNomination.nominee.email,emailContentUpdateDB[0].emailSub,emailContentUpdateDB[0].emailBody,epNominatorsEmail,AllRoles.LA,AllRoles.LA,currentWebUrl,[],currentUserEmail);
                                this.EmailNotification.nominationEmail(body, postURL);
                            }
                            if(intakeNominationData.intakeNomination.nominee.email!='')
                            {
                                const emailContentNominee = await this.EmailNotification.getNotificationList(QCReviewStatus.SubmittedByQC+'-'+NominationStatus.Completed+'-'+'Nominee', { role: AllRoles.QC }, intakeNominationData);
                                if(emailContentNominee && emailContentNominee.length > 0 && emailContentNominee[0].IsEnabled)
                                {
                                    const body = this.makeEmailBody(intakeNominationData.intakeNomination.nominee.email, emailContentNominee[0].emailSub, emailContentNominee[0].emailBody, emailContentNominee[0].emailCC, AllRoles.QC, QCBUTTONSACTIONS.GRANT_STATUS, currentWebUrl, [],currentUserEmail);
                                    this.EmailNotification.nominationEmail(body, postURL);
                                }
                            }                           
                        }
                    }
                

                    
                    //11/3/2021:- Here We are sending email to QC Reviewers with selected Discipline - nominationRevierwsUsers is not blank then send an email and get all the users from the nominationrevierwsusers state 
                    if (isPDStatusGrantedBefore2016) {
                        const emailContentForQC = await this.EmailNotification.getNotificationList(NominationStatus.PendingWithLocalAdmin, { role: AllRoles.NOMINATOR }, intakeNominationData, intakeNominationData.intakeNomination.pdDiscipline);
                        if(emailContentForQC && emailContentForQC.length > 0 && emailContentForQC[0].IsEnabled)
                        {
                            //send an CC email to LA
                            const body= this.makeEmailBody(intakeNominationData.nominationDetailsByLA.assignee.email,emailContentForQC[0].emailSub,emailContentForQC[0].emailBody,intakeNominationData.nominationDetailsByLA.assignee.email,AllRoles.QC,NominationStatus.PendingWithQC,currentWebUrl,[],currentUserEmail);
                            this.EmailNotification.nominationEmail(body, postURL);
                        }
                    }
                    

                    //11/3/2021:- Here We are sending email to QC Reviewers with selected Discipline - nominationRevierwsUsers is not blank then send an email and get all the users from the nominationrevierwsusers state
                    if (isPDStatusGrantedAfter2016) {
                        const emailContentIsLAFalse = await this.EmailNotification.getNotificationList(NominationStatus.PendingWithQC, { role: AllRoles.NOMINATOR }, intakeNominationData, intakeNominationData.intakeNomination.pdDiscipline);
                        if(emailContentIsLAFalse && emailContentIsLAFalse.length > 0 && emailContentIsLAFalse[0].IsEnabled)
                        {
                            const qcDisciplineUsers = nominationReviewersUsers.map((qcusers, i) => { return nominationReviewersUsers[i].AuthorizedQC[0].email; }).join(';');
                            
                            //send an email QC users, cc EP Nominators
                            const body= this.makeEmailBody(qcDisciplineUsers,emailContentIsLAFalse[0].emailSub,emailContentIsLAFalse[0].emailBody,epNominatorsEmail,AllRoles.QC,NominationStatus.PendingWithQC,currentWebUrl,[],currentUserEmail);
                            this.EmailNotification.nominationEmail(body, postURL);
                        }   
                    }
                
                
                }
                else if (action == 2) {
                    actionStatus = saved ? STATUS.DELETE_SUCCESS : STATUS.DELETE_ERROR;
                }
                this.setState({
                    itemDetails: intakeNominationData,
                    intakeNomination: intakeNominationData.intakeNomination,
                    loading: false,
                    actionStatus: actionStatus,
                    isMessageDialogVisible: actionStatus != STATUS.DEFAULT ? true : false
                });
            }
        }
    }

    private  makeEmailBody(To:string,Subject:string,Body:string,CC:string,Actor:string,qcButtonAction:string,WebUrl:string,attachmentURL:string[],currentUser:string)
    {
        CC = CC && CC.length > 0 ? CC : "";
        if(To && Subject && Body && Actor && qcButtonAction){
            const body: string = JSON.stringify({
                'emailaddress': To,
                'emailSubject': Subject,
                'emailBody': Body,
                'emailCC': CC,
                'emailActor': Actor,
                'qcButtonAction':qcButtonAction,
                'emailItemLink': WebUrl,
                'nominationAttachment':attachmentURL,
                'currentUser':this.props.context.pageContext.user.email
            });
            return body;
        }
    }




    private getStatusRelatedMessage(status) {
        switch (status) {
            case STATUS.DELETE_SUCCESS:
                return Messages.DeleteSuccessPrefix;
            case STATUS.DELETE_ERROR:
                return Messages.DeleteFailedPrefix;
            case STATUS.SAVE_SUCCESS:
                return Messages.SaveSuccessPrefix;
            case STATUS.SAVE_ERROR:
                return Messages.SaveFailedPrefix;
            case STATUS.SUBMIT_SUCCESS:
                return Messages.SubmitSuccessPrefix;
            case STATUS.SUBMIT_ERROR:
                return Messages.SubmitFailedPrefix;
        }
    }

    private _getMessage(): JSX.Element {
        let status: string = this.state && this.state.actionStatus;

        if (status == STATUS.DELETE_SUCCESS || status == STATUS.SAVE_SUCCESS || status == STATUS.SUBMIT_SUCCESS) {
            return (

                <Dialog
                    hidden={!this.state.isMessageDialogVisible}
                    dialogContentProps={{
                        onDismiss: this.close, type: DialogType.normal, title: "Message"
                    }}
                    modalProps={dialogModalProps}>
                    <MessageBar

                        messageBarType={MessageBarType.success}
                        isMultiline={true}
                        actions={
                            <div>
                                <MessageBarButton onClick={this.close}>OK</MessageBarButton>
                            </div>
                        }
                    >
                        {this.getStatusRelatedMessage(status)}
                    </MessageBar>
                </Dialog>);
        }
        else if (status == STATUS.DELETE_ERROR || status == STATUS.SAVE_ERROR || status == STATUS.SUBMIT_ERROR) {
            return (
                <Dialog
                    hidden={!this.state.isMessageDialogVisible}
                    dialogContentProps={{
                        type: DialogType.normal, title: "Message"
                    }}
                    modalProps={dialogModalProps}>
                    <MessageBar

                        messageBarType={MessageBarType.error}
                        isMultiline={true}
                        actions={
                            <div>
                                <MessageBarButton onClick={this.close}>OK</MessageBarButton>
                            </div>
                        }>
                        {this.getStatusRelatedMessage(status)}
                    </MessageBar>
                </Dialog>);
        }
    }

}
