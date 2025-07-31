import * as React from 'react';
import *  as NominationLibraryComponent from "pd-nomination-library";
import styles from './Panel.module.scss';
import {addMonths, DatePicker, DefaultButton, defaultDatePickerStrings, Dialog, DialogFooter, DialogType, Dropdown, IDropdownOption, Panel, PanelType, Stack, TextField, Toggle } from '@fluentui/react';
import { IAllNominationDetails, IAttachment, IIntakeNomination, IMasterDetails, INominationDetailsByLA, INominationDetailsByPTPAC, INominationDetailsByQC, INominationReviewer, INomineeDetails, IReferences } from 'pd-nomination-library';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { AllRoles, NominationStatus, PdStatus, QCReviewStatus } from 'pd-nomination-library';
import { INominationListViewItem } from 'pd-nomination-library';
import autobind from 'autobind-decorator';
import SpinnerComponent from '../spinnerComponent/spinnerComponent';
import {GENERICBUTTONSACTIONS, GROUPNAME, PanelPosition, QCBUTTONSACTIONS, stackTokens, STATUS } from '../commonSettings/settings';
import { FileUploader } from '../control/FileUploader';
import { addDays, format } from 'date-fns';
import { ConstantsConfig, IConstants, INITIAL_CANDIDATE_NOMINATED, INITIAL_NOTIFY_OPTIONS, INITIAL_REFERENCESPASSED_AND_QARPASSED_OPTIONS, INITIAL_TRACK_REFERENCES_OPTIONS } from '../models/IUIConstants';
import { IEmployeeUpdateProperties } from 'pd-nomination-library';
import { IProfessionalDesignationDetailed } from 'pd-nomination-library';
import CommonMethods from '../models/CommonMethods';

export interface IQCFormProps {
    position?: PanelPosition;
    onDismiss?: () => void;
    context: WebPartContext;
    invokedItem: INominationListViewItem;
    isNewForm: boolean;
}

export interface IQCFormState {
    isOpen?: boolean;
    isVisible?: boolean;
    isFormStatus?: string;
    isSaveValid: boolean;
    isReferencesValid: {minRequired:number,isRequired:boolean, isHide:boolean};
    isRequestPTPACReviewValid: boolean;
    nomineeDetails?: INomineeDetails;
    itemDetails: IAllNominationDetails;
    intakeNomination: IIntakeNomination;
    pdNominationDetailed: IProfessionalDesignationDetailed[];
    detailsLANomination: INominationDetailsByLA;
    detailsQCNomination: INominationDetailsByQC;
    detailsPTPACNomination: INominationDetailsByPTPAC;
    detailsNominationReferences: IReferences[];
    masterListData: IMasterDetails;
    loading: boolean;
    isSaving: boolean;
    NominationFormAttachment: IAttachment[];
    NominationOtherAttachments: IAttachment[];
    attachmentType: string;
    files: Array<any>;
    showDialog: boolean;
    showWithdrawDialog: boolean;
    showSendEmailDialog: boolean;
    showSendSCForVoteDialog:boolean;
    AddPracticeDirectorInCC: boolean;
    AnyoneElseNotify: string[];
    NominationNotify: string[];
    ReferenceNotify: string[];
    selectedAttachments: string[];
    grantedOn?: string;
    endOn?: Date;
    qcStateAction:string;
    NomineeStatusAlreadyGranted: boolean;
    reviewerUsers: INominationReviewer[];
    attachmentsFolderPath:string;
    existingSubcategory:string[];
}
export default class QualityCoordinatorForm extends React.Component<IQCFormProps, IQCFormState> {
    protected Constants: IConstants = null;
    public masterDetails: IMasterDetails;
    private NominationLibComponent = new NominationLibraryComponent.NominationLibrary(this.props.context);
    private NominationLibMasters = new NominationLibraryComponent.IntakeNominationLibrary(this.props.context);
    private EmailNotification=new NominationLibraryComponent.NotificationList(this.props.context);
    private NominationListLibComponent = new NominationLibraryComponent.NominationListLibrary(this.props.context);
    private readonly currentWebUrl = this.props.context.pageContext.web.absoluteUrl;
    private readonly currentUserEmail = this.props.context.pageContext.user.email;

    private intakeFormDetails = null;
    public constructor(props: IQCFormProps, state: IQCFormState) {
        super(props, state);

        this.state = {
            itemDetails: null,
            intakeNomination: null,
            pdNominationDetailed: null,
            detailsLANomination:null,
            detailsQCNomination:null,
            detailsPTPACNomination:null,
            detailsNominationReferences: null,
            isOpen: true,
            nomineeDetails: null,
            masterListData: null,
            loading: !this.props.isNewForm,
            isSaving: false,
            isReferencesValid:{minRequired:0,isRequired:false,isHide:true},
            selectedAttachments: [],
            NominationFormAttachment: [],
            NominationOtherAttachments: [],
            attachmentType: "Other",
            files: [],
            isSaveValid: false,
            isRequestPTPACReviewValid: false,
            showDialog: false,
            showWithdrawDialog: false,
            showSendEmailDialog:false,
            showSendSCForVoteDialog:false,
            AddPracticeDirectorInCC: false,
            AnyoneElseNotify: [],
            NominationNotify:['Nominee'],
            ReferenceNotify:null,
            grantedOn: CommonMethods.getSPFormatDate(new Date()),
            endOn: new Date(format(addDays(new Date(Date.now()), 365), "MM/dd/yyyy")),
            qcStateAction:'',
            NomineeStatusAlreadyGranted:false,
            reviewerUsers: null,
            attachmentsFolderPath: null,  
            existingSubcategory:[],       
        };
        this.Constants = ConstantsConfig.GetConstants();
    }

   
    private isValidationError(type: number) {
        const {detailsQCNomination, detailsPTPACNomination,NominationFormAttachment, isReferencesValid,intakeNomination, detailsNominationReferences} = this.state;
        let isError = detailsQCNomination && detailsPTPACNomination ? false : true;
        if(detailsQCNomination && detailsPTPACNomination)
        {
            if (!detailsQCNomination.qcStatus && type == 0)
                isError = true;
            if (!detailsQCNomination.reviewNotes && type == 0)
                isError = true;
            if (!detailsPTPACNomination.reviewDueDate && type == 1)
                isError = true;
            if (!intakeNomination.trackCandidateNominated  && type == 1)
                isError = true;    
  
        }
        if(isReferencesValid && type == 1)
        {   
            const notNullReferencesUserLength = detailsNominationReferences && detailsNominationReferences.filter(references => references.referencesUser !== null).length;
            if (detailsNominationReferences && detailsNominationReferences !== undefined && notNullReferencesUserLength < isReferencesValid.minRequired)
                isError = isReferencesValid.isRequired;
        }
        if(intakeNomination && !intakeNomination.billingCode && intakeNomination.pdStatus.toUpperCase() === "SIGNATURE AUTHORITY" && intakeNomination.trackCandidateNominated === INITIAL_CANDIDATE_NOMINATED[2].key && type == 1)
                isError = true; 
        
        return isError;
    }

    private NominationTrackMinLimitAndRequired()
    {
        const { intakeNomination} = this.state;
        let validation = {minRequired: 0, isRequired:false, isHide:true};
        if(intakeNomination && intakeNomination!.pdStatus && intakeNomination!.trackCandidateNominated){
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
            isRequestPTPACReviewValid: this.isValidationError(1),
            isReferencesValid: this.NominationTrackMinLimitAndRequired(),
            loading: false
        });
    }

    public componentDidMount() {
        this.initializeComponent();

    }

    private async initializeComponent() {
        const {intakeNomination} = this.state;
        await this.getMasterDataInformation();
        if (this.props && this.props.invokedItem) {
            await this.getInvokedNominationItem(this.props.invokedItem);
        }

        this.SetValidState();
        
    }
    private async getMasterDataInformation() {
        const masterFormData = await this._handleAsync(this.NominationLibMasters.getMasterDetails());
        this.setState({
            masterListData: masterFormData
        });
    }


    private async getInvokedNominationItem(item: INominationListViewItem) {
        let itemData: IAllNominationDetails = await this._handleAsync(this.NominationLibComponent.getNominationDetails(item.id, item.nominee, { role: AllRoles.QC }));
        
        if(itemData.intakeNomination){
            const nominationReviewersUsers = itemData && itemData.intakeNomination && itemData.intakeNomination.pdDiscipline ?  await this._handleAsync(this.NominationListLibComponent.getQCDisciplineUsers(itemData.intakeNomination.pdDiscipline)) : "";
            const nomineeFormAttach: IAttachment[] = itemData && itemData.nominationAttachments && itemData.nominationAttachments.filter((element) => { return element.attachmentType === "Nomination Form"; });
            const otherAttach: IAttachment[] = itemData && itemData.nominationAttachments && itemData.nominationAttachments.filter((element) => { return element.attachmentType !== "Nomination Form" && element.attachmentType !== null; });

            this.setState(() => ({
                itemDetails: itemData,
                intakeNomination: itemData.intakeNomination,
                detailsLANomination: itemData.nominationDetailsByLA,
                detailsQCNomination: itemData.nominationDetailsByQC,
                detailsPTPACNomination: itemData.nominationDetailsByPTPAC,
                detailsNominationReferences: itemData.nominationReferences,
                NominationFormAttachment: nomineeFormAttach,
                NominationOtherAttachments: otherAttach,
                isFormStatus: itemData.intakeNomination.nominationStatus,
                reviewerUsers: nominationReviewersUsers,
                attachmentsFolderPath: nomineeFormAttach && nomineeFormAttach.length > 0 ? nomineeFormAttach.map(attachment => attachment.attachmentUrl)[0]: null
            }), () => {
                this._handleAsync(this._getNomineeProfessionalDesignation());
                this.checkIfPDStatusWithDisciplineAlreadyGranted();
            });
        }
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

    private checkIfPDStatusWithDisciplineAlreadyGranted(){
        const {pdNominationDetailed, intakeNomination} = this.state;
            if(pdNominationDetailed && pdNominationDetailed.length > 0 && intakeNomination){
            const isExist = pdNominationDetailed.filter(nomineeDetailItem => nomineeDetailItem.friendlyName === intakeNomination.pdDiscipline && nomineeDetailItem.professionalDesignation === intakeNomination.pdStatus).length > 0;
                this.setState({
                    NomineeStatusAlreadyGranted: isExist
                });
            }

    }

    private async _UpdateEmployeeProfessionalDesignation(): Promise<boolean> {
        try {
            let isNomineePDStatusNew: any;
            const {masterListData, intakeNomination, pdNominationDetailed, grantedOn, endOn} = this.state;
            const profDestId =  intakeNomination.pdStatus ? masterListData.professionalDesignation.filter((_pd: any) => _pd._professionalDesignationTitle == intakeNomination.pdStatus)[0]["_professionalDesignationId"] : null;
            const subcategoryId =  intakeNomination.pdSubcategory ? masterListData.pdSubCategory.filter((_sc: any) => intakeNomination.pdSubcategory.indexOf(_sc._pdSub) > -1) : null;
                                                                  
            const disciplineId =  intakeNomination.pdDiscipline ? masterListData.discipline.filter((_disc: any) => _disc._disciplineFriendlyName == intakeNomination.pdDiscipline)[0]["_disciplineId"] : null;
            const proficientLanguagesIds =  intakeNomination.proficientLanguage ? masterListData.language.filter((_lang: any) => intakeNomination.proficientLanguage.indexOf(_lang._langText) > -1) : null;
            

            const nomineeFinanceUserID = parseInt(intakeNomination.financeUserID);
            if(subcategoryId)
            {
                isNomineePDStatusNew = pdNominationDetailed && pdNominationDetailed.length > 0 ? pdNominationDetailed.filter((_disc: any) => _disc.friendlyName == intakeNomination.pdDiscipline && _disc.professionalDesignation == intakeNomination.pdStatus && subcategoryId.map(e => e['_pdSub']).indexOf(_disc.subCategory) === 0) : []; 
            }
            else
            {
                isNomineePDStatusNew = pdNominationDetailed && pdNominationDetailed.length > 0 ? pdNominationDetailed.filter((_disc: any) => _disc.friendlyName == intakeNomination.pdDiscipline && _disc.professionalDesignation == intakeNomination.pdStatus) : []; 

            }
            
            const arrProficientLanguages =  proficientLanguagesIds && proficientLanguagesIds.length > 0 ? proficientLanguagesIds.map(lang => ({ id: null, financeUserId: nomineeFinanceUserID, proficientLaguageId: lang["_langId"], isDelete: false})) : [{id: null, financeUserId: 0, proficientLaguageId: null, isDelete: false }];
            const arrProfessionalDesignation =  profDestId && subcategoryId && subcategoryId.length > 0 ? subcategoryId.map(sub => ({ 
                id: null, 
                financeUserId: nomineeFinanceUserID, 
                designationId: profDestId, 
                pdSubategoryId: sub["_pdId"],
                disciplineId: disciplineId,
                grantedOn: new Date(grantedOn),
                removedOn: intakeNomination.pdStatus === "Limited Signature Authority" ? endOn : null,
                level: null,
                isDelete: false})) 
                : 
                [{ 
                    id: null, 
                    financeUserId: nomineeFinanceUserID, 
                    designationId: profDestId, 
                    pdSubategoryId: subcategoryId,
                    disciplineId: disciplineId,
                    grantedOn: new Date(grantedOn),
                    removedOn: intakeNomination.pdStatus === "Limited Signature Authority" ? endOn : null,
                    level: null,
                    isDelete: false
                }];

            
            const insertEmployeeObject: IEmployeeUpdateProperties = {
                financeUserId: nomineeFinanceUserID,
                committeeAssignments: [{ id: null, financeUserId: 0, committeeId: null, isDelete: false }],
                proficientLanguages: arrProficientLanguages,
                professionalDesignations:arrProfessionalDesignation,
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
    
   

    private dropDownListObject(items: string[]) {
        return items.map(item => { return { "key": item, "text": item }; });
    }

    private onOtherFilesChanged = (files: IAttachment[]) => {
        this.setState({
            NominationOtherAttachments: files
        });
        
    }

    private _getRefer1 = (items: any) => {
      this._getRefereeInfo("referee0", items);
    }
    private _getRefer2 = (items: any) => {
        this._getRefereeInfo("referee1", items);
    }
    private _getRefer3 = (items: any) => {
        this._getRefereeInfo("referee2", items);
    }
    private _getRefer4 = (items: any) => {
        this._getRefereeInfo("referee3", items);
    }
    private _getRefer5 = (items: any) => {
        this._getRefereeInfo("referee4", items);
    }

    private _getRefereeInfo(id: string, items: any) {
        const {detailsNominationReferences} = this.state;
        const indexNumber  = id.substring(id.length, id.length - 1);

        let filterReferences = [];
        
        if(detailsNominationReferences.length > 0)
        {
            filterReferences = detailsNominationReferences.filter(obj1 => obj1.referencesUser !== null && items.find(obj2 => obj1.referencesUser.title === obj2.text));
        }

        if(items.length > 0)
        {
            if(typeof detailsNominationReferences[indexNumber] === 'undefined') {
                detailsNominationReferences.push({id: 0, referencesUser: { title: items[items.length-1].text, id: items[items.length-1].id, email: items[items.length-1].secondaryText }, referencesTrackVal: "Blank"});
            }
            else{
                detailsNominationReferences[indexNumber].referencesUser= { title: items[items.length-1].text, id: items[items.length-1].id, email: items[items.length-1].secondaryText };
            }
        }
        else {
            detailsNominationReferences[indexNumber].referencesUser= null;
        }

        try {
           
            this.setState(() => ({
                detailsNominationReferences: detailsNominationReferences,
                loading: false,
                //ReferenceNotify: filterReferences
            }), () => {
                this.SetValidState();
            });
            
            
        } catch (error) {
            console.log('_getRefereeInformation: ', error);
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

        }
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
            
         }
     }
 
    public initializeIntakeFormPanel() {
       
        const hideDialog: boolean = this.state.showDialog;
        const hideWithdrawDialog: boolean = this.state.showWithdrawDialog;
        const hideSendEmailDialog: boolean = this.state.showSendEmailDialog;
        const hideSendSCForVoteDialog: boolean = this.state.showSendSCForVoteDialog;

        const { isReferencesValid, detailsNominationReferences, intakeNomination, detailsLANomination, detailsPTPACNomination, detailsQCNomination, isFormStatus, masterListData, NominationFormAttachment,NominationOtherAttachments,attachmentsFolderPath, itemDetails } = this.state;
        
        const isDisabled = isFormStatus === NominationStatus.PendingWithQC || isFormStatus === undefined ? false : true;
        const isPTPACDisabled = isFormStatus === NominationStatus.PendingWithPTPACChair || NominationStatus.PendingWithPTPACReviewer || NominationStatus.PendingWithQC || isFormStatus === undefined ? false : true;
        
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
                let userDisplayName = currVal.referencesUser !== null ? "/" + currVal.referencesUser.title : "/";
                prevVal.push([userDisplayName]);
                return prevVal; // *********  Important ******
            }, [])
            : [];


        const ddlProfessionalDesignation = masterListData && masterListData.professionalDesignation ? this.dropDownListObject(masterListData.professionalDesignation.filter((item: any) => item._professionalDesignationTitle !== "Provisional Signature Authority").map((title:any) =>  title._professionalDesignationTitle)) : [];
        const ddlDiscipline = masterListData && masterListData.discipline ? this.dropDownListObject(masterListData.discipline.map((friendlyName:any) =>  friendlyName._disciplineFriendlyName)) : [];
        const ddlLanguage = masterListData && masterListData.language ? this.dropDownListObject(masterListData.language.map((category:any) =>  category._langText)) : [];
        const ddlPDSubCategory = masterListData && masterListData.pdSubCategory ? this.dropDownListSubCategoryObject(masterListData.pdSubCategory.map((category:any) =>  category._pdSub)) : [];
        const ddlReferences = detailsNominationReferences && detailsNominationReferences.length > 0 ? this.dropDownListObject(detailsNominationReferences.filter(filterRefer => filterRefer.referencesUser !== null).map((reference:any) =>  reference.referencesUser.email)) : [];
        const qcAttachmentPath = attachmentsFolderPath && attachmentsFolderPath.substring(0, attachmentsFolderPath.lastIndexOf("/", attachmentsFolderPath.lastIndexOf("/")-1)) + "/QC";

        
        const qcAllOtherAttachments = NominationOtherAttachments && NominationOtherAttachments.length > 0 ?  NominationOtherAttachments.filter(attachment => attachment.attachmentType !== "Nomination Form").map(attachment => {return {"key":qcAttachmentPath + "/" + attachment.attachmentName, "text":attachment.attachmentName};}) : [];
        const removeDuplicateAttachments = qcAllOtherAttachments && qcAllOtherAttachments.length > 0  && qcAllOtherAttachments.filter((s => ({ key }) => !s.has(key) && s.add(key))(new Set));
        
        const _onChangeNominationPasses = (event, checked?: boolean) => {
            const isChecked: boolean = checked;
            this.setState((prevState) => ({
                detailsQCNomination: {
                    ...prevState.detailsQCNomination,
                    nominationPasses: isChecked
                }
            }), () => {
                this.SetValidState();
            });
        };

        const _onChangeAddPracticeDirectorInCC = (event, checked?: boolean) => {
            const isChecked: boolean = checked;
            this.setState((prevState) => ({
                ...prevState,
                AddPracticeDirectorInCC:isChecked
            }), () => {
                this.SetValidState();
            });
        };

        const _onChangeQualityCoordinatorNotes = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newText: string): void => {
            const notesContent: string = newText;
            this.setState((prevState) => ({
                detailsQCNomination: {
                    ...prevState.detailsQCNomination,
                    reviewNotes: notesContent
                }
            }), () => {
                this.SetValidState();
            });

        };
        
        const handlePTPACReviewDueDateChange = (selDate: Date | null | undefined): void => {
            const ptpacReviewDueDate: Date = selDate;
            this.setState((prevState) => ({
                detailsPTPACNomination: {
                    ...prevState.detailsPTPACNomination,
                    reviewDueDate: CommonMethods.getSPFormatDate(ptpacReviewDueDate)
                }, 
              
            }), () => {
                this.SetValidState();
            });

        };
        
        const _onChangeForNominationNotify = (event: any, selectedOption: IDropdownOption,) => {
            const {NominationNotify} = this.state;
            const ddlArray: Array<string> = NominationNotify ? NominationNotify : [];
            if (selectedOption) {
                this.setState(() => ({
                    NominationNotify: selectedOption.selected ? [...ddlArray, selectedOption.key as string] : ddlArray.filter(key => key !== selectedOption.key)
                    
                }), () => {
                    this.SetValidState();
                });
            }
           
        };

        const _onChangeForReferenceNotify = (event: any, selectedOption: IDropdownOption,) => {
            const {ReferenceNotify} = this.state;
            const ddlArray: Array<string> = ReferenceNotify ? ReferenceNotify : [];
            if (selectedOption) {
                this.setState(() => ({
                    ReferenceNotify: selectedOption.selected ? [...ddlArray, selectedOption.key as string] : ddlArray.filter(key => key !== selectedOption.key)
                    
                }), () => {
                    this.SetValidState();
                });
            }
           
        };

        const _onChangeForSendSCForVoteAttachments = (event: any, selectedOption: IDropdownOption,) => {
            const {selectedAttachments} = this.state;
            const ddlArray: any = selectedAttachments ? selectedAttachments : [];
            if (selectedOption) {
                this.setState(() => ({
                    selectedAttachments:  selectedOption.selected ? [...ddlArray, selectedOption.key as string] : ddlArray.filter(key => key !== selectedOption.key)
                }), () => {
                    this.SetValidState();
                });
            }
           
        };


        const _onChangeForTrackNominated = (event, selectedOption: IDropdownOption,) => {
            //const ddlSelectedItem: string = selectedOption.text;
           
            this.setState((prevState) => ({
                intakeNomination: {
                    ...prevState.intakeNomination,
                    [event.target.id]: selectedOption.key as string
                },
            }), () => {
                this.SetValidState();
            });
           
            //  this.validateNotificationPhaseRequiredField(event.target.id); 
        };

        const _onChangeReferencesStatus = (event: any, item: IDropdownOption): void => {
            const itemVal: string = item.text;
            const indexNumber  = event.target.id.substring(event.target.id.length, event.target.id.length - 1);
            detailsNominationReferences[indexNumber].referencesTrackVal = itemVal;
            this.setState(() => ({
                detailsNominationReferences: detailsNominationReferences
            }), () => {
                this.SetValidState();
            });
        };


        const _onChangeReferencesPassed = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
            const itemVal: string = item.text;
            this.setState((prevState) => ({
                detailsQCNomination: {
                    ...prevState.detailsQCNomination,
                    referencesPassed: itemVal
                }
            }), () => {
                this.SetValidState();
            });
        };

        const _onChangeQARPassed = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
            const itemVal: string = item.text;
            this.setState((prevState) => ({
                detailsQCNomination: {
                    ...prevState.detailsQCNomination,
                    qarPassed: itemVal
                }
            }), () => {
                this.SetValidState();
            });
        };

        const _onChangeAddNotify =  (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newText: string): void => {
            const AnyoneElseEmailAddress: string[] = newText.replace(",", ";").split(";");
            this.setState((prevState) => ({
                AnyoneElseNotify:  AnyoneElseEmailAddress
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

        const _onChangeRecommendation = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newText: string): void => {
            const notesContent: string = newText;
            this.setState((prevState) => ({
                detailsPTPACNomination: {
                    ...prevState.detailsPTPACNomination,
                    recommendation: notesContent
                }
            }), () => {
                this.SetValidState();
            });
        };

        const _onChangeForDropDownStringTypeControls = (event, selectedOption: IDropdownOption,) => {
            //const ddlSelectedItem: string = selectedOption.text;
           
            this.setState((prevState) => ({
                intakeNomination: {
                    ...prevState.intakeNomination,
                    [event.target.id]: selectedOption.key as string
                },
            }), () => {
                this.getEmployeeExistingSubCategory();
                this.SetValidState();
            });
            if(event.target.id == "trackCandidateNominated")
            {
                this.clearReferencesValues();
            }
           
            //  this.validateNotificationPhaseRequiredField(event.target.id); 
        };

        const _onChangeEmployeeNumber = (event, checked?: boolean) => {
            const isChecked: boolean = checked;
            this.setState((prevState) => ({
                detailsLANomination: {
                    ...prevState.detailsLANomination,
                    isEmployeeNumberUpdated: isChecked
                }
            }), () => {
                this.SetValidState();
            });
        };

        const _onChangeEmployeeAgreement = (event, checked?: boolean) => {
            const isChecked: boolean = checked;
            this.setState((prevState) => ({
                detailsLANomination: {
                    ...prevState.detailsLANomination,
                    isEmployeeAgreementSigned: isChecked
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
                            <div className={styles.column6}><span className={styles.header}>Quality Coordinator</span> </div>
                        </div>
                       
                        <div className={styles.row}>
                            <div className={styles.column12}>Complete this form to start the Professional Designation nomination process for a candidate. Visit the <a href="https://milliman.sharepoint.com/sites/ProfDesignationNominationSupport/SitePages/QualityCoordinatorForm.aspx" data-interception="off" target="_blank" rel="noopener noreferrer">support site </a> for more information.</div>
                        </div>
                        <div className={[styles.row, styles.rowEnd].join(" , ")}>
                            <div className={styles.column12}><span className={styles.mandatInfo}><strong>Fields marked (<span className={styles.star}>*</span>) are mandatory</strong></span> </div>
                        </div>
                        <div className={styles.row}>
                            <div className={styles.column12 + ' text-label'}><span className={styles.heading}>Candidate</span></div>
                        </div>
                        <div className={styles.row}>
                            <div className={styles.column3 + ' text-label'}>
                                {intakeNomination && intakeNomination.nominee ? <PeoplePicker
                                    context={this.props.context}
                                    titleText="Nominee Name"
                                    personSelectionLimit={1}
                                    showtooltip={true}
                                    //required={true}
                                    disabled={true}
                                    showHiddenInUI={false}
                                    defaultSelectedUsers={intakeNomination && intakeNomination.nominee && intakeNomination.nominee.title ? ["/" + intakeNomination.nominee.title.toString()] : null}
                                    principalTypes={[PrincipalType.User]}
                                    resolveDelay={1000}
                                    ensureUser={true}
                                /> : ""}
                                

                            </div>
                            <div className={styles.column3 + ' text-label'}>
                                <TextField label="Practice"
                                    disabled={true}
                                    //required={true}
                                    value={intakeNomination && intakeNomination.nomineePractice ? intakeNomination.nomineePractice : ""}
                                />
                                {/* {this.state.loading &&
                                <Spinner label='Loading...' ariaLive='assertive' />
                            } */}
                            </div>
                            <div className={styles.column3 + ' text-label'}>
                                <TextField label="Office"
                                    disabled={true}
                                    //required={true}
                                    value={intakeNomination && intakeNomination.nomineeOffice ? intakeNomination.nomineeOffice : ""}
                                />
                            </div>
                            <div className={styles.column3 + ' text-label'}>
                                <TextField label="Discipline"
                                    disabled={true}
                                    //required={true}
                                    value={intakeNomination && intakeNomination.nomineeDiscipline ? intakeNomination.nomineeDiscipline : ""}
                                />
                            </div>
                        </div>
                        <div className={[styles.row, styles.rowEnd].join(" , ")}>
                            <div className={styles.column3 + ' text-label'}>
                                <Toggle
                                    checked={detailsLANomination && detailsLANomination.isEmployeeNumberUpdated ? true : false}
                                    label={
                                        <div>
                                            Employee Number Updated	 
                                        </div>
                                    }
                                    id={"employeeNumber"}
                                    onText="Yes"
                                    offText="No"
                                    onChange={_onChangeEmployeeNumber}
                                    disabled={isPTPACDisabled}
                                />
                            </div>
                            <div className={styles.column3 + ' text-label'}>
                                <Toggle
                                    checked={detailsLANomination && detailsLANomination.isEmployeeAgreementSigned ? true : false}
                                    label={
                                        <div>
                                            Employee Agreements Signed:	 
                                        </div>
                                    }
                                    id={"employeeAgreement"}
                                    onText="Yes"
                                    offText="No"
                                    onChange={_onChangeEmployeeAgreement}

                                    disabled={isPTPACDisabled}
                                />
                            </div>
                            {itemDetails && itemDetails.nominationDetailsByLA && itemDetails.nominationDetailsByLA.assignee ? 
                            <div className={styles.column6 + ' text-label'}>
                                <PeoplePicker
                                    context={this.props.context}
                                    titleText="Nominator(s):"
                                    personSelectionLimit={1}
                                    showtooltip={true}
                                    //required={true}
                                    key={"Assignee"}
                                    disabled={isPTPACDisabled}
                                    ensureUser={true}
                                    showHiddenInUI={false}
                                    principalTypes={[PrincipalType.User]}
                                    defaultSelectedUsers={itemDetails && itemDetails.nominationDetailsByLA && itemDetails.nominationDetailsByLA.assignee ? ["/" + itemDetails.nominationDetailsByLA.assignee.title.toString()] : []}
                                    resolveDelay={1000} />
                            </div>
                             : ""}
                        </div>
                        <div className={styles.row}>
                            <div className={styles.column12 + ' text-label'}><span className={styles.heading}>Nominate For</span></div>
                        </div>
                        <div className={styles.row}>
                            <div className={styles.column3 + ' text-label'}>
                                <Dropdown placeholder="Select Professional Designation"
                                    disabled={isPTPACDisabled}
                                    id={"pdStatus"}
                                    label="Professional Designation"
                                    options={ddlProfessionalDesignation}
                                    //required={true}
                                    onChange={_onChangeForDropDownStringTypeControls}
                                    defaultSelectedKey={intakeNomination && intakeNomination.pdStatus ? intakeNomination.pdStatus : ""}
                                />
                            </div>
                            {intakeNomination && intakeNomination.pdDiscipline === "Employee Benefits" ?
                                <div className={styles.column3 + ' text-label'}>
                                   <Dropdown placeholder="Select Subcategory"
                                        disabled={isPTPACDisabled}
                                        label="Subcategory"
                                        id={"pdSubcategory"}
                                        options={ddlPDSubCategory}
                                        multiSelect
                                        required={true}
                                        onChange={_onChangeForDropDownControls}

                                        defaultSelectedKeys={intakeNomination && intakeNomination.pdSubcategory ? intakeNomination.pdSubcategory : []}
                                    />
                                </div>
                            : ""}
                            <div className={styles.column3 + ' text-label'}>
                                <Dropdown placeholder="Select PD Discipline"
                                    disabled={isPTPACDisabled}
                                    label="PD Discipline"
                                    id={"pdDiscipline"}
                                    options={ddlDiscipline}
                                    //options={this.masterDetails.discipline.length > 0 ? this.masterDetails.discipline.map(stringText => ({key: stringText.abbreviation, text:stringText.friendlyName})): []} 
                                    //required={true}
                                    onChange={_onChangeForDropDownStringTypeControls}
                                    selectedKey={intakeNomination && intakeNomination.pdDiscipline ? intakeNomination.pdDiscipline : ""}
                                />
                            </div>
                            <div className={styles.column3 + ' text-label'}>
                                <PeoplePicker
                                    context={this.props.context}
                                    titleText="EP Nominator(s)"
                                    key={"epNominators"}
                                    personSelectionLimit={5}
                                    showtooltip={true}
                                    //required={true}
                                    disabled={isPTPACDisabled}
                                    ensureUser={true}
                                    showHiddenInUI={false}
                                    onChange={this._getEPNominatorInformation}
                                    principalTypes={[PrincipalType.User]}
                                    defaultSelectedUsers={intakeNomination && intakeNomination.epNominators && intakeNomination.epNominators.length > 0 ? epNominationString : []}
                                    resolveDelay={1000} />
                            </div>
                        </div>
                        <div className={[styles.row, styles.rowEnd].join(" , ")}>
                            <div className={styles.column6 + ' text-label'}>
                                <Dropdown placeholder="Select Language"
                                    disabled={isPTPACDisabled}
                                    id={"proficientLanguage"}
                                    multiSelect
                                    label="List any language in which the candidate is proficient to perform work"
                                    options={ddlLanguage}
                                    onChange={_onChangeForDropDownControls}
                                    defaultSelectedKeys={intakeNomination && intakeNomination.proficientLanguage ? intakeNomination.proficientLanguage : []}

                                />
                            </div>
                            {(intakeNomination && intakeNomination.pdStatus && intakeNomination.trackCandidateNominated && intakeNomination.pdStatus.toUpperCase() === "SIGNATURE AUTHORITY" && intakeNomination.trackCandidateNominated === INITIAL_CANDIDATE_NOMINATED[2].key)?
                                <div className={styles.column3   + ' text-label'}>
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
                        </div>
                        <div className={styles.row}>
                            <div className={styles.column12 + ' text-label'}><span className={styles.heading}>Product Section</span></div>
                            <div className={styles.column12 + ' text-label'}>Enter a deadline for the PTPAC review in the PTPAC Review Due Date field before sending the nomination to PTPAC.</div>
                            
                        </div>
                        
                        <div className={[styles.row, styles.rowEnd].join(" , ")}>
                            {
                                /*
                                <div className={styles.column3 + ' text-label'}>
                                    <Toggle
                                        checked={intakeNomination && intakeNomination.isProductPerson}
                                        label={
                                            <div>
                                                Is Product Person:	 
                                            </div>
                                        }
                                        id={"productPerson"}
                                        onText="Yes"
                                        offText="No"
                                        onChange={_onChangeProductPerson}
                                        disabled={isDisabled}

                                    />
                                </div>
                                */
                            }
                            
                            <div className={styles.column4 + ' text-label'}>
                                <Dropdown placeholder="Select candidate nominated"
                                    disabled={isPTPACDisabled}
                                    label="Under which track is the candidate nominated ?"
                                    id={"trackCandidateNominated"}
                                    options={INITIAL_CANDIDATE_NOMINATED}
                                    required={true}
                                    onChange={
                                        _onChangeForTrackNominated
                                    }
                                    selectedKey={intakeNomination && intakeNomination.trackCandidateNominated ? intakeNomination.trackCandidateNominated : ""}
                                />
                               
                            </div>
                            <div className={styles.column4 + ' text-label'}>
                                <div className={[styles.column12,styles.fontWeightIncrease].join(" , ")}>PTPAC Review Due Date<span className={styles.star}>*</span></div>
                                <DatePicker
                                    isRequired={false}
                                    minDate={addDays(new Date(Date.now()), 0)}
                                    placeholder="Select a date..."
                                    ariaLabel="Select a date"
                                    onSelectDate={handlePTPACReviewDueDateChange}
                                    // DatePicker uses English strings by default. For localized apps, you must override this prop.
                                    //strings={defaultDatePickerStrings}
                                    value={detailsPTPACNomination && detailsPTPACNomination.reviewDueDate ? new Date(format(new Date(detailsPTPACNomination.reviewDueDate.toString()), "MM/dd/yyyy")) : null}
                                    disabled={detailsQCNomination && detailsQCNomination.sentToPTPACDate ? true : false}
                                />
                            </div>
                            <div className={styles.column4 + ' text-label'}>
                                <TextField label="Recommendation"
                                    disabled={isPTPACDisabled}
                                    multiline rows={6}
                                    onChange={_onChangeRecommendation}
                                    value={detailsPTPACNomination && detailsPTPACNomination.recommendation ? detailsPTPACNomination.recommendation : ""}
                                />
                            </div>
                        </div>
                        </div>
                        {
                        isReferencesValid && !isReferencesValid.isHide ?
                        <>
                        <div className={styles.row}>
                            <div className={styles.column12 + ' text-label'}><span className={styles.heading}>References</span></div>
                        </div>
                        <div className={styles.row}>
                            <div className={styles.column12 + ' text-label'}>
                                <label className="text-label ms-Label ms-Dropdown-label root-306" id="Dropdown62-label">
                                A minimum of two references is required for an AP nomination and a minimum of three references is required for a SA/QR nomination. References are not required for a first-time LSA nomination and are encouraged with a renewal LSA nomination.
                                    {
                                        isReferencesValid && isReferencesValid.isRequired?<span className={styles.star}>*</span>:""
                                    }  
                                </label>
                            </div>
                        </div>
                        <div className={styles.row}>
                            <div className={styles.column2 + ' text-label'}>
                                <PeoplePicker
                                    context={this.props.context}
                                    personSelectionLimit={1}
                                    showtooltip={true}
                                    required={detailsNominationReferences && detailsNominationReferences !== undefined && detailsNominationReferences.length < isReferencesValid.minRequired}
                                    key={"referee0"}
                                    disabled={isPTPACDisabled}
                                    onChange={this._getRefer1}
                                    ensureUser={true}
                                    showHiddenInUI={false}
                                    principalTypes={[PrincipalType.User]}
                                    defaultSelectedUsers={detailsNominationReferences && detailsNominationReferences !== undefined && detailsNominationReferences.length >= 1 ? referencesString[0] : []}
                                    resolveDelay={1000} />
                            </div>
                            <div className={styles.column2 + ' text-label'}>
                                <Dropdown placeholder=""
                                    disabled={isPTPACDisabled}
                                    id={"referencesIndex0"}
                                    label=""
                                    options={INITIAL_TRACK_REFERENCES_OPTIONS}
                                    //required={true}
                                    onChange={_onChangeReferencesStatus}
                                    selectedKey={detailsNominationReferences && detailsNominationReferences.length >= 1 ? detailsNominationReferences[0].referencesTrackVal : "Blank"}
                                    //defaultSelectedKey={detailsQCNomination && detailsQCNomination.referencesPassed ? detailsQCNomination.referencesPassed : "No"}
                                />
                            </div>
                        </div>
                        
                        <div className={styles.row}>
                            <div className={styles.column2 + ' text-label'}>
                                <PeoplePicker
                                    context={this.props.context}
                                    personSelectionLimit={1}
                                    showtooltip={true}
                                    required={detailsNominationReferences && detailsNominationReferences !== undefined && detailsNominationReferences.length < isReferencesValid.minRequired}
                                    key={"referee1"}
                                    disabled={isPTPACDisabled}
                                    onChange={this._getRefer2}
                                    ensureUser={true}
                                    showHiddenInUI={false}
                                    principalTypes={[PrincipalType.User]}
                                    defaultSelectedUsers={detailsNominationReferences && detailsNominationReferences !== undefined && detailsNominationReferences.length >= 2 ? referencesString[1] : []}
                                    resolveDelay={1000} />
                            </div>
                            <div className={styles.column2 + ' text-label'}>
                                <Dropdown placeholder=""
                                    disabled={isPTPACDisabled}
                                    id={"referencesIndex1"}
                                    label=""
                                    options={INITIAL_TRACK_REFERENCES_OPTIONS}
                                    //required={true}
                                    onChange={_onChangeReferencesStatus}
                                    selectedKey={detailsNominationReferences && detailsNominationReferences.length >= 2 ? detailsNominationReferences[1].referencesTrackVal : ""}
                                    //defaultSelectedKey={detailsQCNomination && detailsQCNomination.referencesPassed ? detailsQCNomination.referencesPassed : "No"}
                                />
                            </div>
                        </div>

                        <div className={styles.row}>
                            <div className={styles.column2 + ' text-label'}>
                                <PeoplePicker
                                    context={this.props.context}
                                    personSelectionLimit={1}
                                    showtooltip={true}
                                    required={detailsNominationReferences && detailsNominationReferences !== undefined && detailsNominationReferences.length < isReferencesValid.minRequired}
                                    key={"referee2"}
                                    disabled={isPTPACDisabled}
                                    onChange={this._getRefer3}
                                    ensureUser={true}
                                    showHiddenInUI={false}
                                    principalTypes={[PrincipalType.User]}
                                    defaultSelectedUsers={detailsNominationReferences && detailsNominationReferences !== undefined && detailsNominationReferences.length >= 3 ? referencesString[2] : []}
                                    resolveDelay={1000} />
                            </div>
                            <div className={styles.column2 + ' text-label'}>
                                <Dropdown placeholder=""
                                    disabled={isPTPACDisabled}
                                    id={"referencesIndex2"}
                                    label=""
                                    options={INITIAL_TRACK_REFERENCES_OPTIONS}
                                    //required={true}
                                    onChange={_onChangeReferencesStatus}
                                    selectedKey={detailsNominationReferences && detailsNominationReferences.length >= 3 ? detailsNominationReferences[2].referencesTrackVal : ""}
                                    //defaultSelectedKey={detailsQCNomination && detailsQCNomination.referencesPassed ? detailsQCNomination.referencesPassed : "No"}
                                />
                            </div>
                        </div>

                        <div className={styles.row}>
                            <div className={styles.column2 + ' text-label'}>
                                <PeoplePicker
                                    context={this.props.context}
                                    personSelectionLimit={1}
                                    showtooltip={true}
                                    required={detailsNominationReferences && detailsNominationReferences !== undefined && detailsNominationReferences.length < isReferencesValid.minRequired}
                                    key={"referee3"}
                                    disabled={isPTPACDisabled}
                                    onChange={this._getRefer4}
                                    ensureUser={true}
                                    showHiddenInUI={false}
                                    principalTypes={[PrincipalType.User]}
                                    defaultSelectedUsers={detailsNominationReferences && detailsNominationReferences !== undefined && detailsNominationReferences.length >= 4 ? referencesString[3] : []}
                                    resolveDelay={1000} />
                            </div>
                            <div className={styles.column2 + ' text-label'}>
                                <Dropdown placeholder=""
                                    disabled={isPTPACDisabled}
                                    id={"referecesIndex3"}
                                    label=""
                                    options={INITIAL_TRACK_REFERENCES_OPTIONS}
                                    //required={true}
                                    onChange={_onChangeReferencesStatus}
                                    selectedKey={detailsNominationReferences && detailsNominationReferences.length >= 4 ? detailsNominationReferences[3].referencesTrackVal : ""}
                                    //defaultSelectedKey={detailsQCNomination && detailsQCNomination.referencesPassed ? detailsQCNomination.referencesPassed : "No"}
                                />
                            </div>
                        </div>

                        <div className={[styles.row, styles.rowEnd].join(" , ")}>
                            <div className={styles.column2 + ' text-label'}>
                                <PeoplePicker
                                    context={this.props.context}
                                    personSelectionLimit={1}
                                    showtooltip={true}
                                    required={intakeNomination && intakeNomination.references !== undefined && intakeNomination.references.length < isReferencesValid.minRequired}
                                    key={"referee4"}
                                    disabled={isPTPACDisabled}
                                    onChange={this._getRefer5}
                                    ensureUser={true}
                                    showHiddenInUI={false}
                                    principalTypes={[PrincipalType.User]}
                                    defaultSelectedUsers={detailsNominationReferences && detailsNominationReferences !== undefined && detailsNominationReferences.length >= 5 ? referencesString[4] : []}
                                    resolveDelay={1000} />
                            </div>
                            <div className={styles.column2 + ' text-label'}>
                                <Dropdown placeholder=""
                                    disabled={isPTPACDisabled}
                                    id={"referencesIndex4"}
                                    label=""
                                    options={INITIAL_TRACK_REFERENCES_OPTIONS}
                                    //required={true}
                                    onChange={_onChangeReferencesStatus}
                                    selectedKey={detailsNominationReferences && detailsNominationReferences.length >= 5 ? detailsNominationReferences[4].referencesTrackVal : ""}
                                    //defaultSelectedKey={detailsQCNomination && detailsQCNomination.referencesPassed ? detailsQCNomination.referencesPassed : "No"}
                                />
                            </div>

                            <div className={styles.column3 + ' text-label'}>
                                <DefaultButton disabled={detailsNominationReferences && detailsNominationReferences !== undefined && detailsNominationReferences && detailsNominationReferences.filter(references => references.referencesUser !== null).length >= isReferencesValid.minRequired? false: true} onClick={this._showSendEmailDialog.bind(this)} text={QCBUTTONSACTIONS.SEND_EMAIL} />
                            </div>
                        </div>
                        </>
                        : ""}  
                        {intakeNomination && intakeNomination.intakeNotes ?   
                        <>
                        <div className={styles.row}>
                            <div className={styles.column6 + ' text-label'}><span className={styles.headingDisabledLabel}>Comments from Nominator</span></div>
                        </div>
                        <div className={styles.row}>
                            <div className={styles.column12 + ' text-label'}>
                                <TextField label="" multiline rows={6} disabled={isPTPACDisabled}
                                        placeholder=''
                                        id={"intakeNotes"}
                                        value={intakeNomination && intakeNomination.intakeNotes ? intakeNomination.intakeNotes : ""}
                                />
                            </div>    
                        </div>
                        </>
                        : ""}
                        
                        {detailsLANomination && detailsLANomination.reviewNotes ? 
                        <>
                        <div className={styles.row}>
                            <div className={styles.column6 + ' text-label'}><span className={styles.headingDisabledLabel}>Comments from Local Admin</span></div>
                        </div>
                        <div className={styles.row}>
                            <div className={styles.column12 + ' text-label'}>
                                <TextField label="" multiline rows={6} disabled={isPTPACDisabled}
                                        placeholder=''
                                        id={"LocalAdminNotes"}
                                        value={detailsLANomination && detailsLANomination.reviewNotes ? detailsLANomination.reviewNotes  : ""}
                                />
                            </div>
                        </div>
                        </>
                        : ""}
                       
                        {detailsPTPACNomination && detailsPTPACNomination.ptpacChairComments ? 
                        <>
                        <div className={styles.row}>
                            <div className={styles.column6 + ' text-label'}><span className={styles.headingDisabledLabel}>Comments from PTPAC Chair</span></div>
                        </div>
                        <div className={styles.row}>
                            <div className={styles.column12 + ' text-label'}>
                                <TextField label="" multiline rows={6} disabled={isPTPACDisabled}
                                        placeholder=''
                                        id={"LocalAdminNotes"}
                                        value={detailsPTPACNomination && detailsPTPACNomination.ptpacChairComments ? detailsPTPACNomination.ptpacChairComments  : ""}
                                />
                            </div>
                        </div>
                        </>
                        : ""}
                        <div className={styles.row}>
                            <div className={styles.column6 + ' text-label'}><span className={styles.heading}>Private Notes</span></div>
                        </div>
                        <div className={styles.row}>
                            <div className={styles.column3 + ' text-label'}>
                                <Toggle
                                    checked={detailsQCNomination && detailsQCNomination.nominationPasses}
                                    label={
                                        <div>
                                            Nomination passes
                                        </div>
                                    }
                                    id={"nominationpasses"}
                                    onText="Yes"
                                    offText="No"
                                    onChange={_onChangeNominationPasses}
                                    disabled={isPTPACDisabled}

                                />
                            </div>
                            <div className={styles.column3 + ' text-label'}>
                                <Dropdown placeholder="References passed"
                                    disabled={isPTPACDisabled}
                                    id={"referencesPassed "}
                                    label="References passed"
                                    options={INITIAL_REFERENCESPASSED_AND_QARPASSED_OPTIONS}
                                    //required={true}
                                    onChange={_onChangeReferencesPassed}
                                    selectedKey={detailsQCNomination && detailsQCNomination.referencesPassed ? detailsQCNomination.referencesPassed : "No"}
                                    //defaultSelectedKey={detailsQCNomination && detailsQCNomination.referencesPassed ? detailsQCNomination.referencesPassed : "No"}
                                />
                            </div>
                            <div className={styles.column3 + ' text-label'}>
                                <Dropdown placeholder="QAR passed"
                                    disabled={isPTPACDisabled}
                                    id={"qarPassed "}
                                    label="QAR passed"
                                    options={INITIAL_REFERENCESPASSED_AND_QARPASSED_OPTIONS}
                                    onChange={_onChangeQARPassed}
                                    //required={true}
                                    selectedKey={detailsQCNomination && detailsQCNomination.qarPassed ? detailsQCNomination.qarPassed : "No"}
                                    //defaultSelectedKey={detailsQCNomination && detailsQCNomination.qarPassed ? detailsQCNomination.qarPassed : "No"}
                                />
                            </div>
                        </div>
                        <div className={styles.row}>
                            <div className={styles.column12}>Don’t forget to edit your notes before granting status. Once the status has been granted, the notes cannot be edited.<span className={styles.star}>*</span></div>
                        </div>
                        <div className={[styles.row, styles.rowEnd].join(" , ")}>
                            <div className={styles.column12 + ' text-label'}>
                                <TextField label="" multiline rows={6} disabled={isPTPACDisabled}
                                    onChange={_onChangeQualityCoordinatorNotes} placeholder='Enter your notes about this candidate.'
                                    id={"QualityCoordinatorNotes"}
                                    value={detailsQCNomination && detailsQCNomination.reviewNotes ? detailsQCNomination.reviewNotes  : ""}
                                />
                            </div>
                        </div>

                        {intakeNomination && intakeNomination.pdStatus && intakeNomination.pdStatus !== PdStatus.RP ?
                            <div>
                                <div className={styles.row}>
                                    <div className={styles.column6 + ' text-label'}><span className={styles.heading}>Nomination Form<span className={styles.star}>*</span></span></div>
                                </div>
                                <FileUploader
                                    onFilesChanged={(fileItem) => { this.onNominationFilesChanged(fileItem); }}
                                    docType={"Nomination Form"}
                                    context={this.props.context}
                                    disabled={false}
                                    role={AllRoles.QC}
                                    onDocumentDelete={(fileItem) => { this.delQCNominationForm(fileItem); }}
                                    files={NominationFormAttachment && NominationFormAttachment.length > 0 && NominationFormAttachment.map((attachment: IAttachment) => {
                                        return attachment;
                                    })}
                                >
                                </FileUploader>
                           

                            <div className={styles.row}>
                                <div className={styles.column6 + ' text-label'}><span className={styles.heading}>Attachments</span></div>
                            </div>
                            <div className={styles.row}>
                                <div className={styles.column12 + ' text-label'}>
                                <FileUploader
                                    onFilesChanged={(fileItem) => { this.onOtherFilesChanged(fileItem); }}
                                    context={this.props.context}
                                    docType={this.state.attachmentType && this.state.attachmentType}
                                    disabled={false}
                                    role={AllRoles.QC}
                                    onDocumentDelete={!isDisabled ? (fileItem) => { this.delQCOtherAttachments(fileItem);} : null}
                                    files={this.state.NominationOtherAttachments.length > 0 && this.state.NominationOtherAttachments.map((attachment: IAttachment) => {
                                        return attachment;
                                    })}>
                                </FileUploader>
                                </div>
                            </div>
                        </div> : ""}
                        <Dialog
                            isOpen={hideWithdrawDialog}
                            type={DialogType.normal}
                            onDismiss={this._closeWithdrawDialog}
                            title='Withdraw'
                            isBlocking={true}
                            className={'ScriptPart'}
                            >
                            <div>
                                <div>
                                     <span className={styles.heading}><b>Are you sure you want to withdraw this nomination?</b> if you do, it will be archived and no longer editable.</span>
                                </div>
                            </div>
                            <DialogFooter>
                                 <DefaultButton onClick={() => { this.processQualityCoordinatorForm(QCBUTTONSACTIONS.WITHDRAW_NOMINATION);}} text="Yes" />
                                 <DefaultButton onClick={this._closeWithdrawDialog} text="No" />
                            </DialogFooter>
                        </Dialog>


                    
                        <Dialog
                            isOpen={hideDialog}
                            type={DialogType.normal}
                            onDismiss={this._closeGrantDialog}
                            title='Grant Status - details'
                            isBlocking={true}
                            className={'ScriptPart'}
                            >
                            <div>
                                    <div>
                                    {' '}
                                    <DatePicker
                                        isRequired
                                        label="Granted On"
                                        placeholder="Date required with no label..."
                                        ariaLabel="Select a date"
                                        //className={styles.control}
                                        strings={defaultDatePickerStrings}
                                        showWeekNumbers={true}
                                        firstWeekOfYear={1}
                                        onSelectDate={this._onSelectDueDate}
                                        value={this.state.grantedOn ? new Date(this.state.grantedOn):undefined}
                                        />
                                        {intakeNomination && intakeNomination.pdStatus === "Limited Signature Authority" ?
                                     <> 
                                    <DatePicker
                                        disabled
                                        label="End On"
                                        placeholder="Disabled (with label)"
                                        ariaLabel="Disabled (with label)"
                                        strings={defaultDatePickerStrings}
                                        value={this.state.endOn ? this.state.endOn:undefined}
                                    />
                                    </>
                                    : ""}
                                    
                                    <Dropdown placeholder="Nominee"
                                        disabled={isDisabled}
                                        id={"NominationNotify"}
                                        multiSelect
                                        label="Notify"
                                        options={INITIAL_NOTIFY_OPTIONS}
                                        //required={true}
                                        onChange={_onChangeForNominationNotify}
                                        defaultSelectedKeys={this.state.NominationNotify ? this.state.NominationNotify : ["Nominee"]}
                                       
                                    />
                                    
                                    <TextField label="Anyone else"
                                        disabled={false}
                                        placeholder="Use a semicolon between each email address with no spaces"
                                        //required={true}
                                        onChange={_onChangeAddNotify}
                                        value={this.state.AnyoneElseNotify ? this.state.AnyoneElseNotify.join(";") : ""}
                                    />
                                    <Toggle
                                        checked={this.state.AddPracticeDirectorInCC ? true : false}
                                        label={
                                            <div>
                                                Add Practice Director in CC	 
                                            </div>
                                        }
                                        id={"AddPracticeDirectorInCC"}
                                        onText="Yes"
                                        offText="No"
                                        onChange={_onChangeAddPracticeDirectorInCC}
                                        disabled={isDisabled}
                                        inlineLabel

                                    />
                                    </div>
                            </div>
                            <DialogFooter>
                                <DefaultButton onClick={this._closeGrantDialog} text="Cancel" />
                                <DefaultButton onClick={() => { this.processQualityCoordinatorForm(QCBUTTONSACTIONS.GRANT_STATUS); }} text={GENERICBUTTONSACTIONS.SUBMIT} />
                            </DialogFooter>
                        </Dialog>

                        <Dialog
                            isOpen={hideSendEmailDialog}
                            type={DialogType.normal}
                            onDismiss={() => this._closeSendEmailDialog()}
                            title='Send emails to references'
                            isBlocking={true}
                            className={'ScriptPart'}>
                            <div>
                                    <div>
                                        <Dropdown placeholder="References"
                                            disabled={false}
                                            id={"ReferenceNotify"}
                                            multiSelect
                                            label="Reference Notification"
                                            options={ddlReferences}
                                            //required={true}
                                            onChange={_onChangeForReferenceNotify}
                                            defaultSelectedKeys={this.state.ReferenceNotify ? this.state.ReferenceNotify : []}
                                        
                                        />
                                    </div>
                            </div>
                            <DialogFooter>
                                <DefaultButton onClick={() => this._closeSendEmailDialog()} text={GENERICBUTTONSACTIONS.CANCEL} />
                                <DefaultButton onClick={() => { this.processSendEmailToReferences(); }} text={GENERICBUTTONSACTIONS.SEND} />
                            </DialogFooter>
                        </Dialog>

                        <Dialog
                            isOpen={hideSendSCForVoteDialog}
                            type={DialogType.normal}
                            onDismiss={() => this._closeSendSCforVoteDialog()}
                            title='Send documents to SC'
                            isBlocking={true}
                            className={'ScriptPart'}>
                            
                            <div>Select documents(s) to send to the steering committee admin <b>in addition to the nomination form</b><br/><br/></div>
                            <div>
                                    <div>
                                        <Dropdown placeholder="Send attachments"
                                            disabled={isDisabled}
                                            //id={"sendSCForVoteAttachments"}
                                            multiSelect
                                            label="Document(s)"
                                            options={removeDuplicateAttachments ? removeDuplicateAttachments : []}
                                            //required={true}
                                            onChange={_onChangeForSendSCForVoteAttachments}
                                            defaultSelectedKeys={this.state.selectedAttachments ? this.state.selectedAttachments : []}
                                        
                                        />
                                    </div>
                            </div>
                            <DialogFooter>
                                <DefaultButton onClick={() => this._closeSendSCforVoteDialog()} text={GENERICBUTTONSACTIONS.CANCEL} />
                                <DefaultButton onClick={() => {this.processQualityCoordinatorForm(QCBUTTONSACTIONS.SEND_SC_FOR_VOTE); }} text={GENERICBUTTONSACTIONS.SEND} />
                            </DialogFooter>
                        </Dialog>
                </div>
            </React.Fragment>);
    }

    private delQCNominationForm = async (file: string) => {
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

    private delQCOtherAttachments = async (file: string) => {
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


    private onNominationFilesChanged = (files: IAttachment[]) => {
        if(files.length > 0){
            files.forEach(attachment => {
                if (!attachment.attachmentType) {
                attachment.attachmentType = "Nomination Form"; // Replace "Default Type" with your desired value
                }
            });
        }
        this.setState((prevState) => ({
            NominationFormAttachment: files
        }));   
        this.SetValidState();
    }


    /*************************************************************************************
	 * Set Granted Status Date on Modal Dialog
	*************************************************************************************/
    private _onSelectDueDate = (date:Date) => {
        this.setState({
            endOn: new Date(format(addDays(new Date(date.toUTCString()), 365), "MM/dd/yyyy")),
            grantedOn: new Date(format(new Date(date.toUTCString()), "MM/dd/yyyy")).toISOString()
        });
    }

    /*************************************************************************************
	 * Shows the dialog
	 *************************************************************************************/
    private _showGrantDialog() {
		this.setState({ showDialog: true });
	}

    /*************************************************************************************
	 * Close the dialog
	*************************************************************************************/
    private _closeGrantDialog = (ev?: React.MouseEvent<HTMLButtonElement, MouseEvent>) => {
        this.setState({ showDialog: false });
    }
    

    /*************************************************************************************
	 * _handleAsync is async function
	*************************************************************************************/

    private _handleAsync = (promise): Promise<any> => {
        return promise
            .then(data => (data))
            .catch(error => Promise.resolve(error));
    }


    private _clickPanel(e) {
        e.preventDefault();
    }

    //Start:- Confirmation Box//

    private _showWithdrawDialog() {
		this.setState({ showWithdrawDialog: true });
	}

    private _showSendSCforVoteDialog() {
		this.setState({ showSendSCForVoteDialog: true });
	}

    private _closeWithdrawDialog = (ev?: React.MouseEvent<HTMLButtonElement, MouseEvent>) => {
        this.setState({ showWithdrawDialog: false });
    }

    private _closeSendSCforVoteDialog = (ev?: React.MouseEvent<HTMLButtonElement, MouseEvent>) => {
        this.setState({ showSendSCForVoteDialog: false });
    }

    private _showSendEmailDialog(ev?: React.MouseEvent<HTMLButtonElement, MouseEvent>) {
		this.setState({ showSendEmailDialog: true });
	}

    private _closeSendEmailDialog(ev?: React.MouseEvent<HTMLButtonElement, MouseEvent>) {
		this.setState({ showSendEmailDialog: false });
	}


    //End:- Confirmation Box//

    public footerRender() {
        const { isFormStatus, detailsPTPACNomination,NominationFormAttachment,isReferencesValid,detailsNominationReferences } = this.state;
        const isNominationForm = NominationFormAttachment && NominationFormAttachment.length > 0 ? NominationFormAttachment.filter(nominationForm => nominationForm.attachmentType == "Nomination Form") : [];
        const filledReferencesValues =  detailsNominationReferences && detailsNominationReferences.length > 0 ? detailsNominationReferences.filter(ref => ref.referencesUser != null).length : 0;
        const isSCVoteDisabled= isFormStatus !== NominationStatus.ApproveCompleted ? false : true;
        const isWithDrawDisabled =   (isFormStatus !== NominationStatus.ApproveCompleted)  ? false : true;
        const isSaveDisabled =   isFormStatus !== NominationStatus.ApproveCompleted ? false : true;

        const isSendSCforVoteDisabled = (isFormStatus == NominationStatus.PendingWithQC) && isNominationForm.length > 0 ? false : true;
        const isGrantStatusDisabled = isFormStatus == NominationStatus.PendingWithQC &&  isNominationForm.length > 0 && filledReferencesValues >= isReferencesValid.minRequired ? false : true;
        const isRequestPtpacReviewDisabled =  isFormStatus == NominationStatus.PendingWithQC &&  detailsPTPACNomination && detailsPTPACNomination.reviewDueDate && isNominationForm.length > 0 ? false: true ;

        return (
            <div className={styles.footerRow}>
                <div className={styles.column12}>
                    <Stack horizontal tokens={stackTokens}>
                        <DefaultButton onClick={this.close && this.props.onDismiss} disabled={false} text={GENERICBUTTONSACTIONS.CANCEL} />
                        {//<DefaultButton disabled={isSubmitDisabled} onClick={() => { this.processQualityCoordinatorForm(QCBUTTONSACTIONS.REQUESTS_MORE_DETAILS) }} text="Request More Details" />
                        }
                        <DefaultButton disabled={isSaveDisabled} onClick={() => { this.processQualityCoordinatorForm(GENERICBUTTONSACTIONS.SAVE); }}  text={GENERICBUTTONSACTIONS.SAVE} />
                        <DefaultButton disabled={isSaveDisabled} onClick={() => this.processQualityCoordinatorForm(GENERICBUTTONSACTIONS.SAVEANDCLOSE) && this.close && this.props.onDismiss} text={GENERICBUTTONSACTIONS.SAVEANDCLOSE} />
                        <DefaultButton disabled={isWithDrawDisabled} onClick={ this._showWithdrawDialog.bind(this)}  text={QCBUTTONSACTIONS.WITHDRAW_NOMINATION} />
                        <DefaultButton disabled={isSendSCforVoteDisabled} onClick={()=> this._showSendSCforVoteDialog()} text={QCBUTTONSACTIONS.SEND_SC_FOR_VOTE} />
                        <DefaultButton disabled={isGrantStatusDisabled} onClick={ this._showGrantDialog.bind(this)} text={QCBUTTONSACTIONS.GRANT_STATUS} //onClick={() => { this.processQualityCoordinatorForm(QCBUTTONSACTIONS.GRANT_STATUS) }} 
                        />
                        <DefaultButton disabled={isRequestPtpacReviewDisabled} onClick={() => { this.processQualityCoordinatorForm(QCBUTTONSACTIONS.REQUEST_PTPAC_REVIEW); }} text={QCBUTTONSACTIONS.REQUEST_PTPAC_REVIEW} />
                        {
                        //<DefaultButton disabled={isSubmitDisabled} onClick={() => { this.processQualityCoordinatorForm(QCBUTTONSACTIONS.GRANT_ACCESS_TO_SOMEONE) }} text="Grant Access to Someone" />
                        }
                    </Stack>
                </div>
            </div>);
    }




    public render(): React.ReactElement<IQCFormProps> {
    
        return (

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
                    </div>
                </div>
            </Panel >
            // </div >
        );
    }

    @autobind
    private close() {
        //this._onCloseTimer = setTimeout(this._onClose.bind(this), parseFloat(styles.duration));
        this.setState({
            isOpen: false
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

    private async processSendEmailToReferences() {
        const {reviewerUsers,ReferenceNotify, intakeNomination,detailsQCNomination,detailsPTPACNomination,itemDetails, detailsNominationReferences} = this.state;
        const postURL: string = this.Constants.PowerAutomateFlowUrl;
            const filterCondition = { status: "Blank" };

        let allNominationDetails: IAllNominationDetails = {
            ...itemDetails,
            intakeNomination: intakeNomination,
            nominationDetailsByQC: detailsQCNomination,
            nominationDetailsByPTPAC:detailsPTPACNomination
        };

        if(ReferenceNotify && ReferenceNotify.length > 0 && allNominationDetails)
        {

            const emailContent = await this.EmailNotification.getNotificationList(QCBUTTONSACTIONS.SEND_EMAIL.toUpperCase(), { role: AllRoles.QC }, allNominationDetails);
            const qcDisciplineUsers = reviewerUsers.map((ptpacUsers, i) => { return reviewerUsers[i].AuthorizedQC[0].email; }).join(';');

            if(emailContent && emailContent.length > 0 && emailContent[0].IsEnabled && ReferenceNotify.length > 0)
            {
                const body = this.makeEmailBody(ReferenceNotify.join(';'), emailContent[0].emailSub, emailContent[0].emailBody, emailContent[0].emailCC + ";" + qcDisciplineUsers, AllRoles.QC, QCBUTTONSACTIONS.SEND_EMAIL.toUpperCase(), this.currentWebUrl,[],this.currentUserEmail);
                this.EmailNotification.nominationEmail(body, postURL);
            }

            // Update values based on the filter
            const updatedStatusOrReferencesArray = detailsNominationReferences.map(obj => {
                if (obj.referencesTrackVal === filterCondition.status && ReferenceNotify.some(filterEmail => filterEmail === obj.referencesUser.email)) {
                    // Update the desired property
                    return { ...obj, referencesTrackVal: 'Pending' };
                }
                return obj;
            });  
            
            this.setState(() => ({
                detailsNominationReferences: updatedStatusOrReferencesArray,
            }), () => {
                this.SetValidState();
            });
            this._closeSendEmailDialog();
        }
    }
   

    private async processQualityCoordinatorForm(action: string) {
        try
        { 
            let { intakeNomination, detailsNominationReferences, detailsLANomination, detailsQCNomination, detailsPTPACNomination, itemDetails, AnyoneElseNotify, AddPracticeDirectorInCC,grantedOn} = this.state;
            const mergeConcat = (...arrays) => [].concat(...arrays.filter(Array.isArray));
            let updatedAttachments: IAttachment[] = [];
            const notNullReferencesCollection = detailsNominationReferences.filter(refer => refer.id != 0 || refer.referencesUser !== null);

            this.setState({
                loading: true,
                qcStateAction:action.toUpperCase(),
            });

            if (intakeNomination) { 
                if(intakeNomination.nominationStatus)
                {
                    detailsQCNomination = {
                        ...detailsQCNomination,
                    reviewerAssignmentDate: CommonMethods.getSPFormatDate(new Date()),
                    };
                }

                switch(action)
                {
                    case GENERICBUTTONSACTIONS.SAVE:
                    case GENERICBUTTONSACTIONS.SAVEANDCLOSE: {
                        detailsQCNomination = {
                            ...detailsQCNomination,
                        qcStatus:  QCReviewStatus.QCDraft,
                        };
                        break;
                    }

                    case QCBUTTONSACTIONS.REQUEST_PTPAC_REVIEW: 
                    {

                        detailsQCNomination = {
                            ...detailsQCNomination,
                            qcStatus:  QCReviewStatus.SentToPTPAC,
                            sentToPTPACDate: CommonMethods.getSPFormatDate(new Date()),
                        };
                        
                        break;
                    }

                    case QCBUTTONSACTIONS.WITHDRAW_NOMINATION: {
                    
                        detailsLANomination = {
                            ...detailsLANomination,
                            withdrawCompletionDate: CommonMethods.getSPFormatDate(new Date()),
                            
                        };
                        detailsQCNomination = {
                            ...detailsQCNomination,
                            qcStatus:  QCReviewStatus.Withdraw,
                            withdrawnDate: CommonMethods.getSPFormatDate(new Date()),
                            
                        };
                        break;
                    }
                    case QCBUTTONSACTIONS.SEND_SC_FOR_VOTE: {
                    
                        detailsQCNomination = {
                            ...detailsQCNomination,
                            qcStatus:  QCReviewStatus.SentToSCForVote,
                            sentToScDate: CommonMethods.getSPFormatDate(new Date()), 
                        };
                        //TODO: Send Notification to SEND SC for Vote user//
                        break;
                    }

                    case QCBUTTONSACTIONS.GRANT_STATUS: {
                        detailsQCNomination = {
                            ...detailsQCNomination,
                            id:detailsQCNomination && detailsQCNomination.id ? detailsQCNomination.id : 0,
                            qcStatus: QCReviewStatus.SubmittedByQC,
                            granted: grantedOn,
                            endDate: intakeNomination.pdStatus !== PdStatus.RP ? CommonMethods.getSPFormatDate(new Date()): null,
                            reviewDate: CommonMethods.getSPFormatDate(new Date()),
                            notificationRecipient: null,//NominationNotify,
                            anyoneElse: AnyoneElseNotify.join(),
                            addPracticeDirector: AddPracticeDirectorInCC
                        },
                        intakeNomination = {
                            ...intakeNomination,
                            submissionDate: new Date(intakeNomination.submissionDate).toLocaleDateString('en-US'),
                        };
                        break;
                    }
                    
                }

                detailsQCNomination = {
                    ...detailsQCNomination,
                    reviewDate: CommonMethods.getSPFormatDate(new Date()),
                    id:detailsQCNomination && detailsQCNomination.id ? detailsQCNomination.id : 0,
                    reviewer: {id: this.props.context.pageContext.legacyPageContext.userId,title:this.props.context.pageContext.user.displayName, email: this.props.context.pageContext.user.email},
                    //nominationId: itemDetails.intakeNomination.id
                    
                };
                
                
                    
            }
            updatedAttachments = mergeConcat(this.state.NominationFormAttachment, this.state.NominationOtherAttachments);
            let oldAttachments = itemDetails && itemDetails.nominationAttachments;
            updatedAttachments = this.processAttachments(updatedAttachments, oldAttachments);

        
            let allNominationDetails: IAllNominationDetails = {
                ...itemDetails,
                intakeNomination: intakeNomination,
                nominationDetailsByLA: detailsLANomination,
                nominationAttachments: updatedAttachments,
                nominationDetailsByQC: detailsQCNomination,
                nominationDetailsByPTPAC:detailsPTPACNomination,
                nominationReferences: notNullReferencesCollection
            };
            


            await this.processData(allNominationDetails, action);
        }
        catch (err) {
            if (err instanceof Error) {
                console.error(`QC panel- processQualityCoordinatorForm things exploded (${err.message})`);
            }
        }
    }


    private async processData(allNominationData: IAllNominationDetails, action: string) {

    try{
            const {reviewerUsers, NominationNotify, AnyoneElseNotify, AddPracticeDirectorInCC,NominationFormAttachment, NomineeStatusAlreadyGranted, selectedAttachments, detailsNominationReferences}=this.state;
            const postURL: string = this.Constants.PowerAutomateFlowUrl;
            const permissionPostURL: string = this.Constants.PermissionPowerAutomateFlowUrl;

            if (allNominationData.intakeNomination && allNominationData.nominationDetailsByQC) 
            {
                if(!NomineeStatusAlreadyGranted)
                {

                    const notifyNomineeEmail:string= NominationNotify.indexOf(INITIAL_NOTIFY_OPTIONS[0].key.toString()) > -1 ? allNominationData.intakeNomination.nominee.email : '';
                    const notifyNominatorEmail: string = NominationNotify.indexOf(INITIAL_NOTIFY_OPTIONS[1].key.toString()) > -1 ? allNominationData.intakeNomination.nominator.email : '';
                    const notifyEPNominators = NominationNotify.indexOf(INITIAL_NOTIFY_OPTIONS[2].key.toString()) > -1 ? allNominationData.intakeNomination.epNominators.map((element) => { return element.email; }).join(';') : '';
                    const notifyAnyoneElse = AnyoneElseNotify != null ? AnyoneElseNotify.map((anyoneElse) => { return anyoneElse; }).join(';') : '';
                    const notifyPersonals = notifyNomineeEmail + ';' + notifyNominatorEmail + ';' + notifyEPNominators + ';' + notifyAnyoneElse;
                    const emailContent = await this.EmailNotification.getNotificationList(allNominationData.nominationDetailsByQC.qcStatus, { role: AllRoles.QC }, allNominationData, allNominationData.intakeNomination.pdDiscipline,allNominationData.intakeNomination.pdStatus);
                    const emailContentNominee = await this.EmailNotification.getNotificationList(allNominationData.nominationDetailsByQC.qcStatus+'-'+NominationStatus.Completed+'-'+'Nominee', { role: AllRoles.QC }, allNominationData);
                    const emailContentITPeoples = await this.EmailNotification.getNotificationList(allNominationData.nominationDetailsByQC.qcStatus +'-'+ NominationStatus.Completed, { role: AllRoles.QC }, allNominationData);
                    const emailContentForPTPACChair = await this.EmailNotification.getNotificationList(allNominationData.nominationDetailsByQC.qcStatus, { role: AllRoles.QC }, allNominationData);

                    const saved: any = await this.NominationLibComponent.saveNominationDetails(allNominationData, { role: AllRoles.QC }, null, null);
                    if(saved){
                        if (action === QCBUTTONSACTIONS.GRANT_STATUS && allNominationData.nominationDetailsByQC.qcStatus !== QCReviewStatus.Withdraw) {

                            const updateDB = action == QCBUTTONSACTIONS.GRANT_STATUS ? await this._handleAsync(this._UpdateEmployeeProfessionalDesignation()) : "";
                            
                            if (allNominationData.nominationDetailsByQC.qcStatus === QCReviewStatus.SubmittedByQC && typeof emailContent === 'object' && emailContent.length > 0 && emailContent[0].IsEnabled && updateDB) {
                                let emailCC = AddPracticeDirectorInCC ? emailContent[0].emailCC : null;
                                const body = this.makeEmailBody(notifyPersonals, emailContent[0].emailSub, emailContent[0].emailBody, emailCC, AllRoles.QC, QCBUTTONSACTIONS.GRANT_STATUS, this.currentWebUrl, [],this.currentUserEmail);
                                this.EmailNotification.nominationEmail(body, postURL);
                            }
                            if(allNominationData.intakeNomination.nominee.email!='' && emailContentNominee.length > 0 && emailContentNominee[0].IsEnabled && updateDB)
                            {
                                const body = this.makeEmailBody(allNominationData.intakeNomination.nominee.email, emailContentNominee[0].emailSub, emailContentNominee[0].emailBody, emailContentNominee[0].emailCC, AllRoles.QC, QCBUTTONSACTIONS.GRANT_STATUS, this.currentWebUrl, [],this.currentUserEmail);
                                this.EmailNotification.nominationEmail(body, postURL);
                            }
                            if (!allNominationData.intakeNomination.nomineeDesignation && emailContentITPeoples.length > 0 && emailContentITPeoples[0].IsEnabled && updateDB) {
                                //TODO: Send Notification to To: corp.it.help.desk@milliman.com about PD status has been updated in employee Database
                                const body = this.makeEmailBody(emailContentITPeoples[0].emailTo, emailContentITPeoples[0].emailSub, emailContentITPeoples[0].emailBody, emailContentITPeoples[0].emailCC, AllRoles.QC, QCBUTTONSACTIONS.GRANT_STATUS, this.currentWebUrl,[],this.currentUserEmail);
                                this.EmailNotification.nominationEmail(body, postURL);
                            }
                            if (action === QCBUTTONSACTIONS.GRANT_STATUS.toUpperCase() && saved && updateDB) {
                                allNominationData.intakeNomination.nominationStatus = NominationStatus.Completed;
                            }
                        }    
                        if (allNominationData.nominationDetailsByQC.qcStatus !== QCReviewStatus.Withdraw && NominationFormAttachment.length > 0 && NominationFormAttachment[0].attachmentUrl) 
                        {     
                            const permissionParameters = CommonMethods.setPermissionOnAttachment(this.props.context,
                                NominationFormAttachment[0].attachmentUrl.match(/\/Nomination Attachments\/(.*?)\//)[1],
                                AllRoles.PTPAC_CHAIR,
                                this.Constants.SP_LIST_NAMES.NominationDocumentLibraryName,
            
                            ); 
                            this.EmailNotification.nominationAttachmentPermission(permissionParameters, permissionPostURL);

                                
                            if (action === QCBUTTONSACTIONS.SEND_SC_FOR_VOTE && emailContent.length > 0 && emailContent[0].IsEnabled) {
                                let nominationAttachment = NominationFormAttachment.filter(attachment => attachment.attachmentType == "Nomination Form").map((attachment) => "/"+ attachment.attachmentUrl.split('/').splice(3).join("/"));
                                let selectedAttach = selectedAttachments.map((attachment) => "/"+ attachment.split('/').splice(3).join("/"));
                                if(selectedAttachments.length > 0)
                                {
                                    nominationAttachment = [...nominationAttachment, ...selectedAttach];

                                }
                                const body = this.makeEmailBody(emailContent[0].emailTo, emailContent[0].emailSub, emailContent[0].emailBody, emailContent[0].emailCC, AllRoles.QC, QCBUTTONSACTIONS.SEND_SC_FOR_VOTE, this.currentWebUrl.trim(), nominationAttachment,this.currentUserEmail);
                                this.EmailNotification.nominationEmail(body, postURL);
                            }
                            
                            else if (action === QCBUTTONSACTIONS.REQUEST_PTPAC_REVIEW && emailContentForPTPACChair.length > 0 && emailContentForPTPACChair[0].IsEnabled) {
                                const ptpacDisciplineUsers = reviewerUsers.map((ptpacUsers, i) => { return reviewerUsers[i].AuthorizedPTPAC[0].email; }).join(';');

                                const body = this.makeEmailBody(ptpacDisciplineUsers, emailContentForPTPACChair[0].emailSub, emailContentForPTPACChair[0].emailBody, emailContentForPTPACChair[0].emailCC, AllRoles.PTPAC_CHAIR, QCBUTTONSACTIONS.REQUEST_PTPAC_REVIEW.toUpperCase(), this.currentWebUrl, [],this.currentUserEmail);
                                this.EmailNotification.nominationEmail(body, postURL);
                            }       
                        }
                        if(action == GENERICBUTTONSACTIONS.SAVE && allNominationData.nominationDetailsByQC.qcStatus !== QCReviewStatus.Withdraw) {

                            if(allNominationData.nominationDetailsByQC.id == 0 && saved && saved.length > 0 && saved[0].hasOwnProperty('ID')){
                                this.setState((prevState) => ({
                                    detailsQCNomination: {
                                        ...prevState.detailsQCNomination,
                                        id:saved[0]['ID']
                                    }
                                }));
                            }
                            if(saved.responseVal.filter(a => a["odata.type"]))
                            {
                                const trackReference = saved.responseVal.filter(a => a["odata.type"] === "SP.Data.Track_x0020_ReferencesListItem");

                                if(saved && saved.length > 0 && trackReference.length > 0){
                                    const resultTrackReferencesArray = this.updateReferencesArray(detailsNominationReferences, trackReference);
                                    this.setState(() => ({
                                        detailsNominationReferences: resultTrackReferencesArray
                                    }));
                                }
                            }
                        }
                    
                        
                    
                        this.setState({
                            itemDetails: allNominationData,
                            intakeNomination: allNominationData.intakeNomination,
                            detailsLANomination: allNominationData.nominationDetailsByLA,
                            detailsQCNomination: allNominationData.nominationDetailsByQC,
                            detailsPTPACNomination: allNominationData.nominationDetailsByPTPAC,
                            detailsNominationReferences: allNominationData.nominationReferences,
                            loading: false
                        });

                        
                        if(action !== GENERICBUTTONSACTIONS.SAVE)
                        { 
                            this.props.onDismiss();
                        }
                    }
                
                } 
            }
        }
        catch (err) {
            if (err instanceof Error) {
                console.error(`QC panel processData things exploded (${err.message})`);
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

    private updateReferencesArray(original: IReferences[], updated) {
      // Create a map of updated objects for efficient lookup
      const updatedMap: any = new Map(updated.map(obj => [obj.Title, obj]));
    
      // Iterate over the original array
      original.forEach(obj => {
        const updatedObj = updatedMap.get(obj.referencesUser.title);
        if (updatedObj) {
          // Update the properties if an updated object exists
          obj.id = updatedObj.Id;
        }
      });
    
      return original;
    }
}
