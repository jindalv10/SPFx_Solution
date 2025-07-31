import * as React from 'react';
import *  as NominationLibraryComponent from "pd-nomination-library";
import styles from './Panel.module.scss';
import { createTheme, DefaultButton, Dialog, DialogFooter, DialogType, Dropdown, FacepileBase, Panel, PanelType, Stack, TextField, Toggle } from '@fluentui/react';
import { IAllNominationDetails,INominationReviewer, IAttachment, IIntakeNomination, IMasterDetails, INominationDetailsByLA, INomineeDetails } from 'pd-nomination-library';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { AllRoles, NominationStatus, PdStatus, QCReviewStatus } from 'pd-nomination-library';
import { INominationListViewItem } from 'pd-nomination-library';
import autobind from 'autobind-decorator';
import SpinnerComponent from '../spinnerComponent/spinnerComponent';
import {GENERICBUTTONSACTIONS, GROUPNAME, PanelPosition, stackTokens, STATUS } from '../commonSettings/settings';
import { ConstantsConfig, IConstants, INITIAL_CANDIDATE_NOMINATED } from '../models/IUIConstants';
import { IEmployeeUpdateProperties } from 'pd-nomination-library/';
import { IProfessionalDesignationDetailed } from 'pd-nomination-library';
import CommonMethods from '../models/CommonMethods';
export interface ILAAdminFormProps {
    position?: PanelPosition;
    onDismiss?: () => void;
    context: WebPartContext;
    invokedItem: INominationListViewItem;
    isNewForm: boolean;
}

export interface ILAAdminFormState {
    isOpen?: boolean;
    isVisible?: boolean;
    isFormStatus?: string;
    isSaveValid: boolean;
    isSubmitValid: boolean;
    nomineeDetails?: INomineeDetails;
    itemDetails: IAllNominationDetails;
    intakeNomination: IIntakeNomination;
    detailsLANomination: INominationDetailsByLA;
    masterListData: IMasterDetails;
    loading: boolean;
    isSaving: boolean;
    files: Array<any>;
    nominationReviewersUsers: INominationReviewer[];
    pdNominationDetailed: IProfessionalDesignationDetailed[];
    grantedOn?: string;
    errorDialogMessage?: string;
    showErrorDialog?: boolean;

}
export default class LAAdminForm extends React.Component<ILAAdminFormProps, ILAAdminFormState> {

    public masterDetails: IMasterDetails;
    private NominationLibComponent = new NominationLibraryComponent.NominationLibrary(this.props.context);
    private NominationLibMasters = new NominationLibraryComponent.IntakeNominationLibrary(this.props.context);
    private NominationListLibComponent = new NominationLibraryComponent.NominationListLibrary(this.props.context);
    private EmailNotification=new NominationLibraryComponent.NotificationList(this.props.context);
    protected Constants: IConstants = null;
    private intakeFormDetails = null;
    public constructor(props: ILAAdminFormProps, state: ILAAdminFormState) {
        super(props, state);

        this.state = {
            itemDetails: null,
            intakeNomination: null,
            detailsLANomination:null,
            isOpen: true,
            nomineeDetails: null,
            masterListData: null,
            loading: !this.props.isNewForm,
            isSaving: false,
            files: [],
            isSaveValid: false,
            isSubmitValid: false,
            nominationReviewersUsers:null,
            pdNominationDetailed: null,
            grantedOn: CommonMethods.getSPFormatDate(new Date()),
            errorDialogMessage: "",
            showErrorDialog: false,
            
        };
        this.Constants = ConstantsConfig.GetConstants();
    }

   
    private isValidationError(type: number) {
        const {detailsLANomination, itemDetails } = this.state;
        let isError = detailsLANomination ? false : true;
        if (detailsLANomination) {
            if (!detailsLANomination.isEmployeeAgreementSigned)
                isError = true;
            if (!detailsLANomination.isEmployeeNumberUpdated && type == 1)
                isError = true;
            if (!detailsLANomination && detailsLANomination.reviewNotes && type == 1)
                isError = true;
        }
        return isError;
    }
    private SetValidState() {

        this.setState({
            isSaveValid: !this.isValidationError(0),
            isSubmitValid: !this.isValidationError(1),
            loading: false
        });
    }

    public componentDidMount() {
        this.initializeComponent();

    }

    private async initializeComponent() {
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
        let itemData: IAllNominationDetails = await this._handleAsync(this.NominationLibComponent.getNominationDetails(item.id, item.nominee, { role: AllRoles.LA.toUpperCase() }));
        
        this.setState(() => ({
            itemDetails: itemData,
            intakeNomination: itemData.intakeNomination,
            detailsLANomination: itemData.nominationDetailsByLA,
            isFormStatus: itemData.intakeNomination.nominationStatus
        }), () => {
            this._handleAsync(this._getNomineeProfessionalDesignation());
        });
    }

    private dropDownListObject(items: string[]) {
        return items.map(item => { return { "key": item, "text": item }; });
    }

    


    public initializeIntakeFormPanel() {
        const { intakeNomination, detailsLANomination, isFormStatus, masterListData, itemDetails } = this.state;
        const isDisabled = isFormStatus === NominationStatus.PendingWithLocalAdmin || isFormStatus === undefined ? false : true;
        const epNominationString = this.props.invokedItem
            && intakeNomination
            && intakeNomination.epNominators !== undefined
            && intakeNomination.epNominators.length > 0 ?

            intakeNomination.epNominators.reduce((prevVal, currVal: any, idx) => {
                prevVal.push("/" + currVal.title);
                return prevVal; // *********  Important ******
            }, [])
            : [];
            const ddlProfessionalDesignation = masterListData && masterListData.professionalDesignation ? this.dropDownListObject(masterListData.professionalDesignation.filter((item: any) => item._professionalDesignationTitle !== "Provisional Signature Authority").map((title:any) =>  title._professionalDesignationTitle)) : [];
            const ddlDiscipline = masterListData && masterListData.discipline ? this.dropDownListObject(masterListData.discipline.map((friendlyName:any) =>  friendlyName._disciplineFriendlyName)) : [];
            const ddlPDSubCategory = masterListData && masterListData.pdSubCategory ? this.dropDownListObject(masterListData.pdSubCategory.map((category:any) =>  category._pdSub)) : [];
            const ddlLanguage = masterListData && masterListData.language ? this.dropDownListObject(masterListData.language.map((category:any) =>  category._langText)) : [];

            //const isApprovedProfessional_HighestCredentialedProfessional =  intakeNomination && intakeNomination.pdStatus === "Approved Professional" && intakeNomination.trackCandidateNominated === INITIAL_CANDIDATE_NOMINATED[0].key ? false:true;

         
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

        const _onChangeLocalAdminNotes = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newText: string): void => {
            const notesContent: string = newText;
            this.setState((prevState) => ({
                detailsLANomination: {
                    ...prevState.detailsLANomination,
                    reviewNotes: notesContent
                }
            }), () => {
                this.SetValidState();
            });

        };

        const referencesString = this.props.invokedItem
        && intakeNomination
        && intakeNomination.references !== undefined
        && intakeNomination.references.length > 0 ?

        intakeNomination.references.reduce((prevVal, currVal: any, idx) => {
            prevVal.push("/" + currVal.title);
            return prevVal; // *********  Important ******
        }, [])
        : [];


        
        return (
            <React.Fragment>
                {this.state.loading ? <SpinnerComponent text={"Loading..."} /> : ""}
                <div className="ms-Panel-scrollableContent scrollableContent-561" data-is-scrollable="true">
                    <div className="ms-Panel-content content-562">
                        <div className={[styles.row, styles.rowEnd].join(" , ")}>
                            <div className={styles.column6}><span className={styles.header}>Local Admin Form</span> </div>
                        </div>
                        <div className={styles.row}>
                            <div className={styles.column12}>Complete this form to validate that the nominee meets the Recognized Professional status criteria. Visit the <a href="https://milliman.sharepoint.com/sites/ProfDesignationNominationSupport/SitePages/LocalReviewerForm.aspx" data-interception="off" target="_blank" rel="noopener noreferrer">support site</a> for more information.</div>
                        </div>
                        <div className={[styles.row, styles.rowEnd].join(" , ")}>
                            <div className={styles.column12}><span className={styles.mandatInfo}><strong>Fields marked (<span className={styles.star}>*</span>) are mandatory</strong></span> </div>
                        </div>
                        <div className={styles.row}>
                            <div className={styles.column12 + ' text-label'}><span className={styles.heading}>Candidate</span></div>
                        </div>
                        <div className={[styles.row, styles.rowEnd].join(" , ")}>
                            <div className={styles.column3 + ' text-label'}>
                                {intakeNomination && intakeNomination.nominee ? <PeoplePicker
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
                                /> : ""}
                                

                            </div>
                            <div className={styles.column3 + ' text-label'}>
                                <TextField label="Practice"
                                    disabled={true}
                                    required={true}
                                    value={intakeNomination && intakeNomination.nomineePractice ? intakeNomination.nomineePractice : ""}
                                />
                                {/* {this.state.loading &&
                                <Spinner label='Loading...' ariaLive='assertive' />
                            } */}
                            </div>
                            <div className={styles.column3 + ' text-label'}>
                                <TextField label="Office"
                                    disabled={true}
                                    required={true}
                                    value={intakeNomination && intakeNomination.nomineeOffice ? intakeNomination.nomineeOffice : ""}
                                />
                            </div>
                            <div className={styles.column3 + ' text-label'}>
                                <TextField label="Discipline"
                                    disabled={true}
                                    required={true}
                                    value={intakeNomination && intakeNomination.nomineeDiscipline ? intakeNomination.nomineeDiscipline : ""}
                                />
                            </div>
                        </div>

                        <div className={styles.row}>
                            <div className={styles.column12 + ' text-label'}><span className={styles.heading}>Nominate For</span></div>
                        </div>
                        <div className={[styles.row, styles.rowEnd].join(" , ")}>
                            <div className={styles.column3 + ' text-label'}>
                                <Dropdown placeholder="Select Professional Designation"
                                    disabled={true}
                                    id={"pdStatus"}
                                    label="Professional Designation"
                                    options={ddlProfessionalDesignation}
                                    //options={this.masterDetails.professionalDesignation.length > 0 ? this.masterDetails.professionalDesignation.map(stringText => ({key: stringText.code, text:stringText.title})): []} 
                                    required={true}
                                    defaultSelectedKey={intakeNomination && intakeNomination.pdStatus ? intakeNomination.pdStatus : ""}
                                />
                            </div>
                            <div className={styles.column3 + ' text-label'}>
                                <Dropdown placeholder="Select PD Discipline"
                                    disabled={true}
                                    label="PD Discipline"
                                    id={"pdDiscipline"}
                                    options={ddlDiscipline}
                                    //options={this.masterDetails.discipline.length > 0 ? this.masterDetails.discipline.map(stringText => ({key: stringText.abbreviation, text:stringText.friendlyName})): []} 
                                    required={true}
                                    selectedKey={intakeNomination && intakeNomination.pdDiscipline ? intakeNomination.pdDiscipline : ""}
                                />
                            </div>
                            {intakeNomination && intakeNomination.pdDiscipline === "Employee Benefits" ?
                                <div className={styles.column3 + ' text-label'}>
                                    <Dropdown placeholder="Select Subcategory"
                                        disabled={true}
                                        label="Subcategory"
                                        id={"pdSubcategory"}
                                        options={ddlPDSubCategory}
                                        multiSelect
                                        //options={this.masterDetails.pdSubCategory.length > 0 ? this.masterDetails.pdSubCategory.map(stringText => ({key: stringText, text:stringText})): []} 
                                        required={true}
                                        //onChange={_onChangeForDropDownControls
                                            //this.handleChange(e);
                                            //this.intakeNomineeDetails.pdSubcategory = [newValue.text];
                                            //this.validateNotificationPhaseRequiredField(FormField.pdSubcategory);
                                        //}
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
                                    disabled={true}
                                    //onChange={this._getEPNominatorInformation}
                                    ensureUser={true}
                                    showHiddenInUI={false}
                                    principalTypes={[PrincipalType.User]}
                                    //this.state.intakeNomination.epNominators.map(String).join("/").toString()
                                    defaultSelectedUsers={intakeNomination && intakeNomination.epNominators && intakeNomination.epNominators.length > 0 ? epNominationString : []}
                                    resolveDelay={1000} />
                            </div>
                            <div className={styles.column3 + ' text-label'}>
                                <Dropdown placeholder="Select candidate nominated"
                                    disabled={true}
                                    label="Under which track is the candidate nominated ?"
                                    id={"trackCandidateNominate"}
                                    options={INITIAL_CANDIDATE_NOMINATED}
                                    //options={this.masterDetails.discipline.length > 0 ? this.masterDetails.discipline.map(stringText => ({key: stringText.abbreviation, text:stringText.friendlyName})): []} 
                                    //required={true}
                                    //onChange={_onChangeForDropDownStringTypeControls}
                                    selectedKey={intakeNomination && intakeNomination.trackCandidateNominated ? intakeNomination.trackCandidateNominated : ""}
                                />
                            </div>
                        </div>
                        
                        <div className={styles.row}>
                            <div className={styles.column12 + ' text-label'}><span className={styles.heading}>Local Reviewer</span></div>
                        </div>
                        <div className={[styles.row, styles.rowEnd].join(" , ")}>
                            <div className={styles.column12 + ' text-label'}>
                                <label className="text-label ms-Label ms-Dropdown-label root-306" id="Dropdown62-label">
                                    Nominator
                                </label>
                            </div>
                            <div className={styles.column6 + ' text-label'}>
                                <PeoplePicker
                                    context={this.props.context}
                                    personSelectionLimit={1}
                                    showtooltip={true}
                                    required={true}
                                    key={"Assignee"}
                                    disabled={true}
                                    //onChange={this._getNominatorInformation}
                                    ensureUser={true}
                                    showHiddenInUI={false}
                                    principalTypes={[PrincipalType.User]}
                                    defaultSelectedUsers={itemDetails && itemDetails.nominationDetailsByLA && itemDetails.nominationDetailsByLA.assignee ? ["/" + itemDetails.nominationDetailsByLA.assignee.title.toString()] : []}
                                    resolveDelay={1000} />
                            </div>
                        </div>

                        <div className={styles.row}>
                            <div className={styles.column6 + ' text-label'}><span className={styles.heading}>Employee Agreement(s) *</span></div>
                        </div>
                        
                        <div className={styles.row}>
                            <div className={styles.column12 + ' text-label'}>
                                <Toggle
                                    checked={detailsLANomination && detailsLANomination.isEmployeeAgreementSigned ? true : false}
                                    label={
                                        <div>
                                            Confirm that the employee has signed the trade secret and non-solicitation agreements OR that the employee contract agreement(s) contains the restrictive covenants clause in the Milliman Employee Portal (UKG/UltiPro). *							
                                        </div>
                                    }
                                    id={"employeeAgreement"}
                                    inlineLabel
                                    onText="Yes"
                                    offText="No"
                                    onChange={_onChangeEmployeeAgreement}
                                    disabled={isDisabled}
                                />
                            </div>
                        </div>
                        <div className={styles.row}>
                            <div className={styles.column6 + ' text-label'}><span className={styles.heading}>Employee Number *</span></div>
                        </div>
                        {intakeNomination && detailsLANomination ?
                            <div className={[styles.row, styles.rowEnd].join(" , ")}>
                                <div className={styles.column12 + ' text-label'}>
                                    <Toggle
                                        checked={detailsLANomination && detailsLANomination.isEmployeeNumberUpdated ? true : false}
                                        label={
                                            <div>
                                               Confirm that the nominee's employee number’s 5th digit has been updated to an 8 (xxxx8xxx) (pg 62-66 of the <a href="https://milliman.sharepoint.com/:b:/r/sites/GCSAcctFin/AccountingManualAttachments/TimeSheet Admin UG 052307.pdf?csf=1&web=1" data-interception="off" target="_blank" rel="noopener noreferrer">Timesheet Admin manual</a>).	 
                                            </div>
                                        }
                                        id={"employeeNumber"}
                                        inlineLabel
                                        onText="Yes"
                                        offText="No"
                                        onChange={_onChangeEmployeeNumber}
                                        disabled={isDisabled}

                                    />
                                </div>
                            </div>
                        : ""}
                        {intakeNomination && intakeNomination.pdStatus && intakeNomination.pdStatus !== PdStatus.RP ?
                            <>
                            <div className={styles.row}>
                                <div className={styles.column6 + ' text-label'}><span className={styles.heading}>Comments</span></div>
                            </div>
                            <div className={[styles.row, styles.rowEnd].join(" , ")}>
                                <div className={styles.column12 + ' text-label'}>
                                    <TextField label="" multiline rows={6} disabled={isDisabled}
                                        onChange={_onChangeLocalAdminNotes} placeholder='Enter comments to your discipline quality coordinator '
                                        id={"LocalAdminNotes"}
                                        value={detailsLANomination && detailsLANomination.reviewNotes ? detailsLANomination.reviewNotes  : ""}
                                    />
                                </div>
                            </div>
                            </>
                            :""
                        }
                            
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
        const { isFormStatus, isSubmitValid } = this.state;
        const isSubmitDisabled = isSubmitValid && isFormStatus === NominationStatus.PendingWithLocalAdmin ? false : true;

        return (
            <div className={styles.footerRow}>
                <div className={styles.column12}>
                    <Stack horizontal tokens={stackTokens}>
                        <DefaultButton onClick={this.close && this.props.onDismiss} disabled={false} text={GENERICBUTTONSACTIONS.CANCEL} />
                        <DefaultButton disabled={isSubmitDisabled} onClick={() => { this.processLocalAdminForm(false); }} text={GENERICBUTTONSACTIONS.SUBMIT} />
                    </Stack>
                </div>
            </div>);
    }




    public render(): React.ReactElement<ILAAdminFormProps> {

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
                        {this.state.showErrorDialog &&

                        <Dialog
                            hidden={!this.state.showErrorDialog} // Only show when there’s an error
                            onDismiss={() => this.setState({ showErrorDialog: false })}
                            dialogContentProps={{
                                type: DialogType.normal,
                                title: 'Error Occurred',
                                subText: this.state.errorDialogMessage
                            }}
                            modalProps={{
                                isBlocking: false
                            }}
                            >
                            <DialogFooter>
                                <DefaultButton onClick={() => this.setState({ showErrorDialog: false })} text="OK" />
                            </DialogFooter>
                        </Dialog>
                        }
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
            const subategoryId =  intakeNomination.pdSubcategory ? masterListData.pdSubCategory.filter((_sc: any) => _sc._pdSub == intakeNomination.pdSubcategory)[0]["_pdId"] : null;
            const discipineId =  intakeNomination.pdDiscipline ? masterListData.discipline.filter((_disc: any) => _disc._disciplineFriendlyName == intakeNomination.pdDiscipline)[0]["_disciplineId"] : null;
            const proficientLanguagesIds =  intakeNomination.proficientLanguage ? masterListData.language.filter((_lang: any) => intakeNomination.proficientLanguage.indexOf(_lang._langText) > -1) : null;
            

            const nomineeFinanceUserID = parseInt(intakeNomination.financeUserID);

            const isNomineePDStatusNew = pdNominationDetailed && pdNominationDetailed.length > 0 ? pdNominationDetailed.filter((_disc: any) => _disc.friendlyName == intakeNomination.pdDiscipline && _disc.professionalDesignation == intakeNomination.pdStatus) : []; 
            const arrProficientLanguages =  proficientLanguagesIds && proficientLanguagesIds.length > 0 ? proficientLanguagesIds.map(lang => ({ id: null, financeUserId: nomineeFinanceUserID, proficientLaguageId: lang["_langId"], isDelete: false})) : [{id: null, financeUserId: 0, proficientLaguageId: null, isDelete: false }];

            
            const insertEmployeeObject: IEmployeeUpdateProperties = {
                financeUserId: nomineeFinanceUserID,
                committeeAssignments: [{ id: null, financeUserId: 0, committeeId: null, isDelete: false }],
                proficientLanguages: arrProficientLanguages,
                professionalDesignations: [{
                        id: null,
                        financeUserId: nomineeFinanceUserID,
                        designationId: profDestId,
                        pdSubategoryId: subategoryId,
                        disciplineId: discipineId,
                        grantedOn: new Date(grantedOn),
                        removedOn: null,
                        level: null,
                        isDelete: false
                }],
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

   

    private async processLocalAdminForm(isDraft: boolean) {
        let { intakeNomination, detailsLANomination, itemDetails } = this.state;
        let qcUsers = null;
        if(intakeNomination.pdDiscipline)
        {
            qcUsers = await this._handleAsync(this.NominationListLibComponent.getQCDisciplineUsers(intakeNomination.pdDiscipline));
        }  
        this.setState({
            loading: true,
            nominationReviewersUsers: qcUsers
        });
       
        if (intakeNomination && detailsLANomination) {
        
            intakeNomination = {
                ...intakeNomination,
                nominationStatus: NominationStatus.SubmittedByLocalAdmin
                //nominationStatus: NominationStatus.PendingWithQC
            };
            
            detailsLANomination = {
                ...detailsLANomination,
                reviewDate: CommonMethods.getSPFormatDate(new Date()),
                //employeeNumberReversedDate: new Date(new Date().toLocaleString()),
                //id:itemDetails.nominationDetailsByLA.id,
                //title:itemDetails.nominationDetailsByLA.title,
                //nominationId: itemDetails.nominationDetailsByLA.nominationId
            };
        }
       
        let allNominationDetails: IAllNominationDetails = {
            ...itemDetails,
            intakeNomination: intakeNomination,
            nominationDetailsByLA: detailsLANomination,
        };

        await this.processData(allNominationDetails, "SUBMIT");
    }

    

    private async processData(allNominationData: IAllNominationDetails, action: string) {
        const{nominationReviewersUsers}=this.state;

        // Ensure 'intakeNomination' data is available to proceed

        if (allNominationData.intakeNomination) {
            let saved = false;
            try {
                // Attempt to save the nomination details using NominationLibComponent
                saved = await this.NominationLibComponent.saveNominationDetails(
                  allNominationData,
                  { role: AllRoles.LA.toUpperCase() },
                  null,
                  null
                );
              } catch (err) {
                // If saving fails, log the error and display an error message in the dialog
                this.setState({
                  errorDialogMessage: 'Failed to save nomination details. Please try again.',
                  showErrorDialog: true,
                  loading: false
                });
                return; // Exit the function as we encountered an error
            }

            // Check if we should send an email based on employee details being updated\
            const isSendAnEmail : boolean = allNominationData.nominationDetailsByLA.isEmployeeNumberUpdated && allNominationData.nominationDetailsByLA.isEmployeeAgreementSigned;
            if(saved)
            {

               // Continue only if saving is successful

                // Extract PD status and other necessary flags
                const rpSelected = allNominationData.intakeNomination.pdStatus.indexOf(PdStatus.RP) > -1;
                const withoutRPSelected = allNominationData.intakeNomination.pdStatus.indexOf(PdStatus.RP) === -1;
                const currentUserEmail = this.props.context.pageContext.user.email;
                const postURL: string = this.Constants.PowerAutomateFlowUrl;
                const currentWebUrl = this.props.context.pageContext.web.absoluteUrl;

                
                // Here We have to Notify to Nominee for Selected RP 
                if(rpSelected && action == "SUBMIT" && allNominationData.nominationDetailsByLA.assignee.email!=null)
                {
                    try {
                        // Check PD status and update employee professional designation if needed
                        if (
                            allNominationData.intakeNomination.pdStatus === PdStatus.RP &&
                            !allNominationData.intakeNomination.isStatusGrantedAfter2016
                        ) {
                            await this._handleAsync(this._UpdateEmployeeProfessionalDesignation());
                        }
                    } catch (err) {
                        // If updating PD fails, display the error message in the dialog
                        console.error("Error updating PD:", err);
                        this.setState({
                          errorDialogMessage: 'Unable to update professional designation. Please contact support.',
                          showErrorDialog: true, // Show the error dialog
                          loading: false
                        });
                        return; // Exit the function after error handling
                      }
                    
                    const emailContent = await this.EmailNotification.getNotificationList(NominationStatus.PendingWithLocalAdmin+'-'+PdStatus.RP, {role: AllRoles.LA}, allNominationData,allNominationData.intakeNomination.pdDiscipline);
                    if(emailContent && emailContent.length > 0 && emailContent[0].IsEnabled)
                    {
                        if (typeof emailContent === 'object' && emailContent !== null && isSendAnEmail) {
                            const body= this.makeEmailBody(allNominationData.intakeNomination.nominee.email,emailContent[0].emailSub,emailContent[0].emailBody,allNominationData.nominationDetailsByLA.assignee.email,AllRoles.LA,AllRoles.LA,currentWebUrl,[],currentUserEmail);
                            this.EmailNotification.nominationEmail(body, postURL);
                        }
                    }
                    
                    
                    const emailContentITPeoples = await this.EmailNotification.getNotificationList(QCReviewStatus.SubmittedByQC+'-'+ NominationStatus.Completed, { role: AllRoles.QC }, allNominationData);
                    if(emailContentITPeoples && emailContentITPeoples.length > 0 && emailContentITPeoples[0].IsEnabled)
                    {
                        const body = this.makeEmailBody(emailContentITPeoples[0].emailTo, emailContentITPeoples[0].emailSub, emailContentITPeoples[0].emailBody, emailContentITPeoples[0].emailCC, AllRoles.QC, 'GRANT_STATUS', currentWebUrl, [],currentUserEmail);
                        this.EmailNotification.nominationEmail(body, postURL);
                    }
                }
                 // Here We have to Notify the QC Reviwers with selected discipline once the Local Admin Submit the form //
                if(allNominationData.nominationDetailsByLA.assignee.email!=null && withoutRPSelected)
                {
                    const emailContent = await this.EmailNotification.getNotificationList(NominationStatus.PendingWithQC, { role: AllRoles.NOMINATOR }, allNominationData);
                    if(emailContent && emailContent.length > 0 && emailContent[0].IsEnabled)
                    {
                        const qcDisplineUsers = nominationReviewersUsers.map((qcusers,i)=>{return nominationReviewersUsers[i].AuthorizedQC[0].email;}).join(';');
                        const body= this.makeEmailBody(qcDisplineUsers,emailContent[0].emailSub,emailContent[0].emailBody,emailContent[0].emailCC,AllRoles.QC,NominationStatus.PendingWithQC,currentWebUrl,[],currentUserEmail);
                        this.EmailNotification.nominationEmail(body, postURL);   
                    }     
                }
              
             
                this.setState({
                    itemDetails: allNominationData,
                    intakeNomination: allNominationData.intakeNomination,
                    detailsLANomination:allNominationData.nominationDetailsByLA, 
                    loading: false
                });
            } else {
                // If saving the nomination failed
                this.setState({
                errorDialogMessage: 'Saving nomination failed. Please check your data and try again.',
                showErrorDialog: true, // Show the error dialog
                loading: false
                });
            }
            this.props.onDismiss();
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
                'currentUser':currentUser
            });
            return body;
        }
    }

}
