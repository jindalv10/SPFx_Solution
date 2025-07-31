import * as React from 'react';
import *  as NominationLibraryComponent from "pd-nomination-library";
import styles from './Panel.module.scss';
import {  DatePicker, DefaultButton, defaultDatePickerStrings, Dialog, DialogFooter, DialogType, Dropdown, IComboBoxOption, IDropdownOption, Panel, PanelType, PrimaryButton, Stack, TextField, Toggle } from '@fluentui/react';
import { IAllNominationDetails, IAttachment, IIntakeNomination, IMasterDetails, INominationDetailsByLA, INominationDetailsByPTPAC, INominationDetailsByQC, INominationReviewer, INomineeDetails } from 'pd-nomination-library';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { AllRoles, NominationStatus, PdStatus, QCReviewStatus, ReviewStatus } from 'pd-nomination-library';
import { INominationListViewItem } from 'pd-nomination-library';
import autobind from 'autobind-decorator';
import SpinnerComponent from '../spinnerComponent/spinnerComponent';
import {GENERICBUTTONSACTIONS, PanelPosition, PTPACBUTTONSACTIONS, stackTokens, STATUS } from '../commonSettings/settings';
import { FileUploader } from '../control/FileUploader';
import { addDays, format } from 'date-fns';
import { ConstantsConfig, IConstants, INITIAL_CANDIDATE_NOMINATED, INITIAL_NOTIFY_OPTIONS } from '../models/IUIConstants';
import { IProfessionalDesignationDetailed } from 'pd-nomination-library';
import CommonMethods from '../models/CommonMethods';

export interface IPTPACFormProps {
    position?: PanelPosition;
    onDismiss?: () => void;
    context: WebPartContext;
    invokedItem: INominationListViewItem;
    isNewForm: boolean;
}

export interface IPTPACFormState {
    isOpen?: boolean;
    isVisible?: boolean;
    isFormStatus?: string;
    isRequestPTPACReview: boolean;
    nomineeDetails?: INomineeDetails;
    itemDetails: IAllNominationDetails;
    intakeNomination: IIntakeNomination;
    pdNominationDetailed: IProfessionalDesignationDetailed[];
    detailsLANomination: INominationDetailsByLA;
    detailsQCNomination: INominationDetailsByQC;
    detailsPTPACNomination: INominationDetailsByPTPAC;
    masterListData: IMasterDetails;
    loading: boolean;
    NominationFormAttachment: IAttachment[];
    NominationOtherAttachments: IAttachment[];
    attachmentType: string;
    files: Array<any>;
    showDialog: boolean;
    internalReviewDueDate?: Date;
    nominationReviewersUsers: INominationReviewer[];

}
export default class PTPACCHAIRFORM extends React.Component<IPTPACFormProps, IPTPACFormState> {

    public masterDetails: IMasterDetails;
    private NominationLibComponent = new NominationLibraryComponent.NominationLibrary(this.props.context);
    private NominationListLibComponent = new NominationLibraryComponent.NominationListLibrary(this.props.context);
    private NominationLibMasters = new NominationLibraryComponent.IntakeNominationLibrary(this.props.context);
    private EmailNotification=new NominationLibraryComponent.NotificationList(this.props.context);
    private readonly currentWebUrl = this.props.context.pageContext.web.absoluteUrl;
    private readonly currentUserEmail = this.props.context.pageContext.user.email;

    protected Constants: IConstants = null;
    public constructor(props: IPTPACFormProps, state: IPTPACFormState) {
        super(props, state);

        this.state = {
            itemDetails: null,
            intakeNomination: null,
            pdNominationDetailed: null,
            detailsLANomination:null,
            detailsQCNomination:null,
            detailsPTPACNomination:null,
            isOpen: true,
            nomineeDetails: null,
            masterListData: null,
            loading: !this.props.isNewForm,
            NominationFormAttachment: [],
            NominationOtherAttachments: [],
            attachmentType: "Other",
            files: [],
            isRequestPTPACReview: false,
            showDialog: false,         
            internalReviewDueDate: new Date(Date.now()),
            nominationReviewersUsers: null,

            
        };
        this.Constants = ConstantsConfig.GetConstants();
    }

   
    private isValidationError(type: number) {
        const {detailsPTPACNomination, intakeNomination} = this.state;
        let isError = detailsPTPACNomination ? false : true;
        if (detailsPTPACNomination) {
            if (!detailsPTPACNomination.recommendation)
                isError = true;
        }
        if(intakeNomination && !intakeNomination.billingCode && intakeNomination.pdStatus.toUpperCase() === "SIGNATURE AUTHORITY" && intakeNomination.trackCandidateNominated === INITIAL_CANDIDATE_NOMINATED[2].key && type == 1)
            isError = true; 
        return isError;
    }
    private SetValidState() {

        this.setState({
            isRequestPTPACReview: !this.isValidationError(1),
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
        let itemData: IAllNominationDetails = await this._handleAsync(this.NominationLibComponent.getNominationDetails(item.id, item.nominee, { role: AllRoles.PTPAC_CHAIR }));
        if(itemData){
            const nomineeFormAttach: IAttachment[] = itemData && itemData.nominationAttachments && itemData.nominationAttachments.filter((element) => { return element.attachmentType === "Nomination Form"; });
            const otherAttach: IAttachment[] = itemData && itemData.nominationAttachments && itemData.nominationAttachments.filter((element) => { return element.attachmentType !== "Nomination Form" && element.attachmentType !== null; });
        
            this.setState({
                itemDetails: itemData,
                intakeNomination: itemData.intakeNomination,
                detailsLANomination: itemData.nominationDetailsByLA,
                detailsQCNomination: itemData.nominationDetailsByQC,
                detailsPTPACNomination: itemData.nominationDetailsByPTPAC,
                NominationFormAttachment: nomineeFormAttach,
                NominationOtherAttachments: otherAttach,
                isFormStatus: itemData.intakeNomination.nominationStatus,
                
            });
        }

    }

    private dropDownListObject(items: string[]) {
        return items.map(item => { return { "key": item, "text": item }; });
    }

    private onPTPACChairFilesChanged = (files: IAttachment[]) => {
        this.setState({
            NominationOtherAttachments: files
        });
    }
    


    public initializeIntakeFormPanel() {
        //const today =format(new Date(),'dd.MM.yyyy');
       
        const hideDialog: boolean = this.state.showDialog;
        const { intakeNomination, detailsPTPACNomination, isFormStatus, masterListData, NominationFormAttachment, itemDetails } = this.state;
   
        const isDisabled = isFormStatus === NominationStatus.PendingWithPTPACChair || isFormStatus === undefined ? false : true;
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
         
           


       

        const _onChangePTPACChairNotes = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newText: string): void => {
            const notesContent: string = newText;
            this.setState((prevState) => ({
                detailsPTPACNomination: {
                    ...prevState.detailsPTPACNomination,
                    ptpacChairComments: notesContent
                }
            }), () => {
                this.SetValidState();
            });

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


        return (
            <React.Fragment>
                {this.state.loading ? <SpinnerComponent text={"Loading..."} /> : ""}
                <div className="ms-Panel-scrollableContent scrollableContent-561" data-is-scrollable="true">
                    <div className="ms-Panel-content content-562">
                        <div className={[styles.row, styles.rowEnd].join(" , ")}>
                            <div className={styles.column6}><span className={styles.header}>PTPAC Chair</span> </div>
                        </div>
                        <div className={styles.row}>
                            <div className={styles.column12}>Assign a PTPAC reviewer and review their recommendation before sending it back to the Discipline Quality Coordinator. Visit the <a href="https://milliman.sharepoint.com/sites/ProfDesignationNominationSupport/SitePages/QualityCoordinatorForm.aspx" data-interception="off" target="_blank" rel="noopener noreferrer">support site </a> for more information.</div>
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
                        {itemDetails && itemDetails.nominationDetailsByLA && itemDetails.nominationDetailsByLA.assignee && 

                            <div className={styles.column6 + ' text-label'}>
                                <PeoplePicker
                                    context={this.props.context}
                                    titleText="Nominator(s)"
                                    personSelectionLimit={1}
                                    showtooltip={true}
                                    //required={true}
                                    key={"Assignee"}
                                    disabled={true}
                                    ensureUser={true}
                                    showHiddenInUI={false}
                                    principalTypes={[PrincipalType.User]}
                                    defaultSelectedUsers={itemDetails && itemDetails.nominationDetailsByLA && itemDetails.nominationDetailsByLA.assignee ? ["/" + itemDetails.nominationDetailsByLA.assignee.title.toString()] : []}
                                    resolveDelay={1000} />
                            </div>
                        }   
                        </div>
                        

                        <div className={styles.row}>
                            <div className={styles.column12 + ' text-label'}><span className={styles.heading}>Nominate For</span></div>
                        </div>
                        <div className={styles.row}>
                            <div className={styles.column3 + ' text-label'}>
                                <Dropdown placeholder="Select Professional Designation"
                                    disabled={true}
                                    id={"pdStatus"}
                                    label="Professional Designation"
                                    options={ddlProfessionalDesignation}
                                    //required={true}
                                    defaultSelectedKey={intakeNomination && intakeNomination.pdStatus ? intakeNomination.pdStatus : ""}
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
                                        //required={true}
                                        defaultSelectedKeys={intakeNomination && intakeNomination.pdSubcategory ? intakeNomination.pdSubcategory : []}
                                    />
                                </div>
                            : ""}
                            <div className={styles.column3 + ' text-label'}>
                                <Dropdown placeholder="Select PD Discipline"
                                    disabled={true}
                                    label="PD Discipline"
                                    id={"pdDiscipline"}
                                    options={ddlDiscipline}
                                    //options={this.masterDetails.discipline.length > 0 ? this.masterDetails.discipline.map(stringText => ({key: stringText.abbreviation, text:stringText.friendlyName})): []} 
                                    //required={true}
                                    //onChange={_onChangeForDropDownStringTypeControls}
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
                                    disabled={true}
                                    ensureUser={true}
                                    showHiddenInUI={false}
                                    principalTypes={[PrincipalType.User]}
                                    defaultSelectedUsers={intakeNomination && intakeNomination.epNominators && intakeNomination.epNominators.length > 0 ? epNominationString : []}
                                    resolveDelay={1000} />
                            </div>
                        </div>
                        <div className={[styles.row, styles.rowEnd].join(" , ")}>
                            <div className={styles.column6 + ' text-label'}>
                                <Dropdown placeholder="Select Language"
                                    disabled={true}
                                    id={"proficientLanguage"}
                                    multiSelect
                                    label="List any language in which the candidate is proficient to perform work"
                                    options={ddlLanguage}
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
                            <div className={styles.column6 + ' text-label'}><span className={styles.heading}>Product Section</span></div>
                        </div>
                        <div className={[styles.row, styles.rowEnd].join(" , ")}>
                            <div className={styles.column4 + ' text-label'}>
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
                            <div className={styles.column3 + ' text-label'}>
                                <DatePicker
                                    isRequired={true}
                                    minDate={addDays(new Date(Date.now()), 0)}
                                    label="PTPAC Chair Review Due Date"
                                    placeholder="Select a date..."
                                    ariaLabel="Select a date"
                                    // DatePicker uses English strings by default. For localized apps, you must override this prop.
                                    strings={defaultDatePickerStrings}
                                    value={detailsPTPACNomination && detailsPTPACNomination.reviewDueDate ? new Date(format(new Date(detailsPTPACNomination.reviewDueDate.toString()), "MM/dd/yyyy")) : null}
                                    disabled={true}

                                />
                            </div>
                            <div className={styles.column3 + ' text-label'}>
                                <TextField label="Recommendation"
                                    disabled={false}
                                    required={true}
                                    multiline rows={6}
                                    onChange={_onChangeRecommendation}
                                    value={detailsPTPACNomination && detailsPTPACNomination.recommendation ? detailsPTPACNomination.recommendation : ""}
                                />
                            </div>
                        </div> 

                        <div className={styles.row}>
                            <div className={styles.column6 + ' text-label'}><span className={styles.heading}>Notes</span></div>
                        </div>
                        <div className={[styles.row, styles.rowEnd].join(" , ")}>
                            <div className={styles.column12 + ' text-label'}>
                                <TextField label="" multiline rows={6} disabled={false}
                                    onChange={_onChangePTPACChairNotes} placeholder='Enter your comments'
                                    id={"ptpacChairNotes"}
                                    value={detailsPTPACNomination && detailsPTPACNomination.ptpacChairComments ? detailsPTPACNomination.ptpacChairComments  : ""}
                                />
                            </div>
                        </div>

                        {intakeNomination && intakeNomination.pdStatus && intakeNomination.pdStatus !== PdStatus.RP ?
                            <div>
                                <div className={styles.row}>
                                    <div className={styles.column6 + ' text-label'}><span className={styles.heading}>Nomination Form </span></div>
                                </div>
                                <FileUploader
                                    onFilesChanged={(fileItem) => { this.onNominationFilesChanged(fileItem); }}
                                    docType={this.state.attachmentType && this.state.attachmentType}
                                    context={this.props.context}
                                    disabled={true}
                                    role={AllRoles.PTPAC_CHAIR}
                                    files={NominationFormAttachment && NominationFormAttachment.length > 0 && NominationFormAttachment.map((attachment: IAttachment) => {
                                        return attachment;
                                    })}
                                >
                                </FileUploader>
                                <div className={styles.row}>
                                    <div className={styles.column6 + ' text-label'}><span className={styles.heading}>Attachments</span></div>
                                </div>
                                <FileUploader
                                    onFilesChanged={(fileItem) => { this.onPTPACChairFilesChanged(fileItem); }}
                                    context={this.props.context}
                                    docType={"Other"}
                                    disabled={isDisabled}
                                    role={AllRoles.PTPAC_CHAIR}
                                    files={this.state.NominationOtherAttachments.length > 0 && this.state.NominationOtherAttachments.map((attachment: IAttachment) => {
                                        return attachment;
                                    })}
                                >
                                </FileUploader>
                            </div> : ""
                        }

                        
                    
                        <Dialog
                            isOpen={hideDialog}
                            type={DialogType.normal}
                            onDismiss={this._closeGrantDialog}
                            title='Assign Reviewer- details'
                            isBlocking={true}
                            className={'ScriptPart'}
                            >
                            <div>
                              <div>
                                <PeoplePicker
                                            titleText='Assigned To'
                                            context={this.props.context}
                                            personSelectionLimit={1}
                                            showtooltip={true}
                                            required={true}
                                            key={"Reviewer"}
                                            disabled={false}
                                            onChange={this._getReviewerInformation}
                                            ensureUser={true}
                                            showHiddenInUI={false}
                                            principalTypes={[PrincipalType.User]}
                                            resolveDelay={1000}
                                />
                                <DatePicker
                                        label="Review Due Date"
                                        placeholder="Date required with no label..."
                                        ariaLabel="Select a date"
                                        minDate={addDays(new Date(Date.now()), 0)}
                                        strings={defaultDatePickerStrings}
                                        showWeekNumbers={true}
                                        firstWeekOfYear={1}
                                        onSelectDate={this._onSelectDueDate}
                                        value={this.state.internalReviewDueDate ? this.state.internalReviewDueDate:undefined}
                                />
                              </div>
                            </div>
                            <DialogFooter>
                                <DefaultButton onClick={this._closeGrantDialog} text= {GENERICBUTTONSACTIONS.CANCEL} />
                                <DefaultButton disabled={this.state.detailsPTPACNomination && this.state.detailsPTPACNomination.reviewer && this.state.detailsPTPACNomination.reviewer.title ? false : true} onClick={() => { this.processPTPACChairForm(GENERICBUTTONSACTIONS.SUBMIT); }} text={GENERICBUTTONSACTIONS.SUBMIT} />
                            </DialogFooter>
                        </Dialog>
                    </div>
                </div>
            </React.Fragment>);
    }


    private _getReviewerInformation = (items: any[]) => {
        if(items.length > 0)
        {
            this.setState((prevState) => ({
                detailsPTPACNomination: {
                    ...prevState.detailsPTPACNomination,
                    reviewer: { title: items[0].text, id: items[0].id, email: items[0].secondaryText},
                }
            }));
        }
        else{
            this.setState((prevState) => ({
                detailsPTPACNomination: {
                    ...prevState.detailsPTPACNomination,
                    reviewer: null,
                }
            }));
        }
    }


    private onNominationFilesChanged = (files: IAttachment[]) => {
        this.setState((prevState) => ({
            NominationFormAttachment: files
        }));   
        this.SetValidState();
    }

    /*************************************************************************************
	 * Set PTPAC Reviewer Internal Due Date on Modal Dialog
	*************************************************************************************/
    private _onSelectDueDate = (date:Date) => {
        this.setState({internalReviewDueDate: new Date(format(new Date(date.toString()), "MM/dd/yyyy"))});
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
        this.setState((prevState) => ({
            detailsPTPACNomination: {
                ...prevState.detailsPTPACNomination,
                reviewer: null,
            },
            showDialog: false
        }));
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

   

    public footerRender() {
        const { isFormStatus,detailsPTPACNomination,isRequestPTPACReview} = this.state;
        const isSendToQCDisabled = isRequestPTPACReview && (isFormStatus === NominationStatus.PendingWithPTPACChair || !isFormStatus) ? false : true;
        const isAssignPTPACReviewerDisabled = isFormStatus === NominationStatus.PendingWithPTPACChair && detailsPTPACNomination && detailsPTPACNomination.recommendationSentDate === null ? false : true;
        return (
            <div className={styles.footerRow}>
                <div className={styles.column12}>
                    <Stack horizontal tokens={stackTokens}>
                        <DefaultButton onClick={this.close && this.props.onDismiss} disabled={false} text="Cancel" />
                        <DefaultButton  disabled={isSendToQCDisabled} onClick={() => { this.processPTPACChairForm(PTPACBUTTONSACTIONS.SEND_TO_QC); }} text={PTPACBUTTONSACTIONS.SEND_TO_QC} />
                        <DefaultButton  disabled={isAssignPTPACReviewerDisabled} onClick={ this._showGrantDialog.bind(this)} text={PTPACBUTTONSACTIONS.ASSIGN_A_PTPAC_REVIEWER} />
                    </Stack>
                </div>
            </div>);
    }




    public render(): React.ReactElement<IPTPACFormProps> {
    
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
   

    private async processPTPACChairForm(action: string) {
        let { intakeNomination, detailsQCNomination, detailsPTPACNomination, itemDetails, internalReviewDueDate} = this.state;
        const mergeConcat = (...arrays) => [].concat(...arrays.filter(Array.isArray));
        let updatedAttachments: IAttachment[] = [];
        let qcUsers = null;


        if(intakeNomination.pdDiscipline)
        {
            qcUsers = await this._handleAsync(this.NominationListLibComponent.getQCDisciplineUsers(intakeNomination.pdDiscipline));
        }  
        this.setState({
            loading: true,
            nominationReviewersUsers: qcUsers
        });

        if (intakeNomination) {


            switch(action)
            {
                case GENERICBUTTONSACTIONS.SUBMIT: {
                    detailsPTPACNomination = {
                        ...detailsPTPACNomination,
                        id:detailsPTPACNomination && detailsPTPACNomination.id ? detailsPTPACNomination.id : 0,
                        reviewerAssignmentDate:CommonMethods.getSPFormatDate(new Date()),
                        internalReviewDueDate: new Date(internalReviewDueDate.toString()).toDateString(),
                        ptpacChair:{id: this.props.context.pageContext.legacyPageContext.userId,title:this.props.context.pageContext.user.displayName, email: this.props.context.pageContext.user.email}
                    };   
                    break;
                }
                
                case PTPACBUTTONSACTIONS.SEND_TO_QC: {
                    detailsQCNomination = {
                        ...detailsQCNomination,
                        qcStatus: ReviewStatus.SubmittedByPTPACChair,
                    };
                    detailsPTPACNomination = {
                        ...detailsPTPACNomination,
                        id:detailsPTPACNomination && detailsPTPACNomination.id ? detailsPTPACNomination.id : 0,
                        reviewDate:CommonMethods.getSPFormatDate(new Date())

                    };
                    break;
                }
                
            }
        
        }
        updatedAttachments = mergeConcat(this.state.NominationFormAttachment, this.state.NominationOtherAttachments);
        let oldAttachments = itemDetails && itemDetails.nominationAttachments;
        updatedAttachments = this.processAttachments(updatedAttachments, oldAttachments);

       
        let allNominationDetails: IAllNominationDetails = {
            ...itemDetails,
            intakeNomination: intakeNomination,
            nominationAttachments: updatedAttachments,
            nominationDetailsByQC: detailsQCNomination,
            nominationDetailsByPTPAC:detailsPTPACNomination
        };

        await this.processData(allNominationDetails, action);
    }

    private async processData(allNominationData: IAllNominationDetails, action: string) {
        const{nominationReviewersUsers, NominationFormAttachment}=this.state;

        if (allNominationData.intakeNomination) {
            const postURL: string = this.Constants.PowerAutomateFlowUrl;
            const permissionPostURL: string = this.Constants.PermissionPowerAutomateFlowUrl;

            const saved = await this.NominationLibComponent.saveNominationDetails(allNominationData, { role: AllRoles.PTPAC_CHAIR }, null, null);
            if(saved)
            {
                const emailContentForPTPACReviewer = await this.EmailNotification.getNotificationList(NominationStatus.PendingWithPTPACReviewer, { role: AllRoles.PTPAC_CHAIR }, allNominationData);
                const emailContentForSendToQC = await this.EmailNotification.getNotificationList(allNominationData.nominationDetailsByQC.qcStatus, { role: AllRoles.PTPAC_CHAIR }, allNominationData);
                const qcDisciplineUsers = nominationReviewersUsers.map((qcUsers, i) => { return nominationReviewersUsers[i].AuthorizedQC[0].email; }).join(';');

                
                if (action === GENERICBUTTONSACTIONS.SUBMIT && emailContentForPTPACReviewer.length > 0 && emailContentForPTPACReviewer[0].IsEnabled) {
                    const body = this.makeEmailBody(allNominationData.nominationDetailsByPTPAC.reviewer.email, emailContentForPTPACReviewer[0].emailSub, emailContentForPTPACReviewer[0].emailBody, emailContentForPTPACReviewer[0].emailCC, AllRoles.PTPAC_REVIEWER, NominationStatus.PendingWithPTPACReviewer, this.currentWebUrl,[],this.currentUserEmail);
                    this.EmailNotification.nominationEmail(body, postURL);
                }
                if (action === PTPACBUTTONSACTIONS.SEND_TO_QC && emailContentForSendToQC.length > 0 && emailContentForSendToQC[0].IsEnabled) {
                    const body = this.makeEmailBody(qcDisciplineUsers, emailContentForSendToQC[0].emailSub, emailContentForSendToQC[0].emailBody, emailContentForSendToQC[0].emailCC, AllRoles.QC, NominationStatus.PendingWithQC, this.currentWebUrl, [],this.currentUserEmail);
                    this.EmailNotification.nominationEmail(body, postURL);
                }
                if(allNominationData.nominationDetailsByPTPAC && allNominationData.nominationDetailsByPTPAC.reviewer &&  allNominationData.nominationDetailsByPTPAC.reviewer.email){
                    const permissionParameters = CommonMethods.setPermissionOnAttachment(this.props.context,
                        NominationFormAttachment[0].attachmentUrl.match(/\/Nomination Attachments\/(.*?)\//)[1],
                        AllRoles.PTPAC_CHAIR,
                        this.Constants.SP_LIST_NAMES.NominationDocumentLibraryName,
                        allNominationData.nominationDetailsByPTPAC.reviewer.email

                    ); 
                    this.EmailNotification.nominationAttachmentPermission(permissionParameters, permissionPostURL);
                }
        
                saved ? console.info(STATUS.SAVE_SUCCESS) : console.error(STATUS.SAVE_ERROR);
                this.setState({
                    itemDetails: allNominationData,
                    intakeNomination: allNominationData.intakeNomination,
                    detailsLANomination: allNominationData.nominationDetailsByLA,
                    detailsQCNomination: allNominationData.nominationDetailsByQC,
                    detailsPTPACNomination: allNominationData.nominationDetailsByPTPAC,
                    loading: false
                });
            }
            else{
                
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
                'currentUser':this.props.context.pageContext.user.email,
            });
            return body;
        }
    }
}
