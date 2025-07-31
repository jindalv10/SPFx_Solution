import * as React from 'react';
import styles from './Nomination.module.scss';
import { INominationProps } from './INominationProps';
import {
  CommandBar,
  DetailsList,
  IColumn,
  ICommandBarItemProps,
  IDetailsList,

} from '@fluentui/react';

import IntakeFormPanel from './Forms/IntakePanel';
import { IDetailsListGroupedNominationState } from './IDetailsListGroupedNominationState';
import *  as NominationLibraryComponent from "pd-nomination-library";
import LAAdminForm from './Forms/LAAdminForm';
import { AllRoles, NominationStatus } from 'pd-nomination-library';
import { INominationListViewItem } from 'pd-nomination-library';
import { IDetailsListGroupedNominationItem } from './IDetailsListGroupedNominationItem';
import autobind from 'autobind-decorator';
import { commandBarStyles, NominationListColumns, OtherDetailsListColumns, PanelPosition, QCDetailsListColumns, PTPACChairDetailsListColumns, PTPACReviewerDetailsListColumns } from './commonSettings/settings';
import QualityCoordinatorForm from './Forms/QCPanel';
import { format } from 'date-fns';
import PTPACCHAIRFORM from './Forms/PTPAChair';
import PTPACReviwerForm from './Forms/PTPACReviewer';
import ErrorBoundary from './models/ErrorBoundary';


export default class Nomination extends React.Component<INominationProps, IDetailsListGroupedNominationState> {

  private _root = React.createRef<IDetailsList>();

  /***************************************************************************
   * Library  Component Service used to perform REST calls
   ***************************************************************************/
  private NominationQueryService = new NominationLibraryComponent.NominationListLibrary(this.props.context);



  constructor(props: INominationProps) {
    super(props);
    this.state = {
      pendingItems: [],
      completedItems:[],
      masterItems: [],
      isOpen: false,
      selectedItem: null,
      actor: null,
      isNew: false,
      columns: null
    };
    // this.NominationDetailsList = NominationListColumns;
  }

  /*************************************************************************************
   * Loads the external scritps sequentially (one after the other) if any
   *************************************************************************************/
   


  /*************************************************************************************
   * Called once after initial rendering
   *************************************************************************************/
  public componentDidMount(): void {
    //const userDetails: IUserDetails = {role: this.props.formType};
    //https://millimandev.sharepoint.com/teams/DeveloperSite-MLT/SitePages/Nominee-Forms.aspx?Actor=Nominator&loadSPFX=true&debugManifestsFile=https%3A%2F%2Flocalhost%3A4321%2Ftemp%2Fmanifests.js

    let params = (new URL(window.location.href)).searchParams;
    // const formType = params.get('FormType'); // is the string "NewForm".
    const actor = params.get('Actor'); // is the string "NewForm".
    const panel = params.get('Panel'); // is the string "NewForm".
    if(panel) {
      this.setState({
        isOpen: true,
        selectedItem: null,
        isNew: true
      });
    }
    
    this.initializeNominationsList(actor.toUpperCase());
  }

  private async initializeNominationsList(actor: string) {
    let allColumns = OtherDetailsListColumns;
    let pendingViewItems =  null, completedViewItems = null, viewItems = null;
    allColumns = actor.toUpperCase() === AllRoles.NOMINATOR.toUpperCase() ? NominationListColumns : allColumns;
    allColumns = actor.toUpperCase() === AllRoles.LA.toUpperCase() ? OtherDetailsListColumns : allColumns;
    allColumns = actor.toUpperCase() === AllRoles.QC.toUpperCase() ? QCDetailsListColumns : allColumns;
    allColumns = actor.toUpperCase() === AllRoles.PTPAC_CHAIR.toUpperCase() ? PTPACChairDetailsListColumns : allColumns;
    allColumns = actor.toUpperCase() === AllRoles.PTPAC_REVIEWER.toUpperCase() ? PTPACReviewerDetailsListColumns : allColumns;
    

    allColumns.map((e) => {
      e.onColumnClick = this._onColumnClick;
    });
    if (actor) {
      const listItems = await this.getNominationListItems(actor);
      if (listItems)
        viewItems = this.getNominationListViewItems(listItems);

      if(actor !== null) {
          switch (actor.toUpperCase()) {
            case AllRoles.NOMINATOR.toUpperCase():
            case  AllRoles.LA.toUpperCase():
                pendingViewItems = viewItems;
                break;
            case AllRoles.QC.toUpperCase():
                pendingViewItems = viewItems ? viewItems.filter((elem) => {return elem.InternalStatus == NominationStatus.PendingWithQC || elem.InternalStatus == NominationStatus.PendingWithPTPACReviewer || elem.InternalStatus == NominationStatus.PendingWithPTPACChair;}) :  "";
                completedViewItems = viewItems ? viewItems.filter((elem) => {return elem.InternalStatus == NominationStatus.ApproveCompleted || elem.InternalStatus == NominationStatus.WithdrawnCompleted;}) :  [];
                break;
            case AllRoles.PTPAC_CHAIR.toUpperCase():
              pendingViewItems = viewItems ? viewItems.filter((elem) => {return elem.InternalStatus == NominationStatus.PendingWithPTPACReviewer || elem.InternalStatus == NominationStatus.PendingWithPTPACChair; }) :  "";
              completedViewItems = viewItems ? viewItems.filter((elem) => {return elem.InternalStatus == NominationStatus.WithdrawnCompleted || elem.InternalStatus == NominationStatus.ApproveCompleted || elem.InternalStatus == NominationStatus.PendingWithQC; }) :  [];
              break;
            case AllRoles.PTPAC_REVIEWER.toUpperCase():
                pendingViewItems = viewItems ? viewItems.filter((elem) => {return elem.InternalStatus == NominationStatus.PendingWithPTPACReviewer; }) : ""  ;
                break;    
            }
      
        this.setState({
          pendingItems: pendingViewItems,
          completedItems: completedViewItems,
          masterItems: listItems,
          actor: actor,
          selectedItem: null,
          columns: allColumns
        });
      }
    }
  }

  private getNominationListViewItems(listItems: INominationListViewItem[]) {
    let viewItems = null;
    if (listItems) {
      viewItems = listItems.map((item) => {
        return {
          key: item.id,
          Nominee: item.nominee ? item.nominee.title : "",
          InternalStatus: item.status,
          PDStatus: item.pdStatus,
          PDDiscipline: item.pdDiscipline,
          Nominator: item.nominator.title,
          EPnominator: item.epNominators ? item.epNominators.map((eachItem: any) => { return eachItem.title; }).join(", ") : "",
          Submitted: item.submitted ? format(new Date(item.submitted.toString())  , "MM/dd/yyyy") : "",
          SendSCforVoteDate:item.sendSCforVoteDate ? format(new Date(item.sendSCforVoteDate.toString())  , "MM/dd/yyyy") : "", //format(item.sendSCforVoteDate, "MM/dd/yyyy")
          NominationPasses:item.nominationPasses && item.nominationPasses.toString() ? "Yes" : "No" ,
          ReferencesPassed:item.referencesPassed ? item.referencesPassed.toString() : "No" ,
          QARPassed:item.qarPassed ? item.qarPassed.toString() : "No",
          PTPACDueDate: item.PTPACDueDate ? format(new Date(item.PTPACDueDate.toString())  , "MM/dd/yyyy") : "", //format(item.sendSCforVoteDate, "MM/dd/yyyy")
          PTPACInternalDueDate: item.PTPACInternalDueDate ? format(new Date(item.PTPACInternalDueDate.toString())  , "MM/dd/yyyy") : "", //format(item.sendSCforVoteDate, "MM/dd/yyyy")
          PTPACReviewer: item.PTPACReviewer && item.PTPACReviewer.title,
          Subcategory: item.Subcategory ? item.Subcategory.map((eachItem: any) => { return eachItem ; }).join(", ") : ""

        };
      });
    }
    return viewItems;
  }


  /**************************************************************************************************
   * Returns a sorted array of all available items for the specified context
   **************************************************************************************************/
  private async getNominationListItems(actor) {
    let [nominationsListItems, nominationsListErr] = await this._handleAsync(this.NominationQueryService.getNominationList({ role: actor }));
    return nominationsListItems ? nominationsListItems : null;
  }

  private _handleAsync = (promise): Promise<any> => {
    return promise
      .then(data => ([data, undefined]))
      .catch(error => Promise.resolve([undefined, error]));
  }

  private formPanel() {
    if (this.state && this.state.isOpen && this.state.actor) {
      if (this.state.selectedItem) {
        switch (this.state.actor.toUpperCase()) {
          case AllRoles.NOMINATOR.toUpperCase(): //Intake Form
            return <IntakeFormPanel isNewForm={false} invokedItem={this.state.selectedItem} context={this.props.context} position={PanelPosition.Right} onDismiss={this.onPanelClosed.bind(this)}></IntakeFormPanel>;
          case AllRoles.LA.toUpperCase(): //LA Admin Form
            return <LAAdminForm isNewForm={false} invokedItem={this.state.selectedItem} context={this.props.context} position={PanelPosition.Right} onDismiss={this.onPanelClosed.bind(this)}></LAAdminForm>;
          case AllRoles.QC.toUpperCase(): //QC Form
            return <QualityCoordinatorForm isNewForm={false} invokedItem={this.state.selectedItem} context={this.props.context} position={PanelPosition.Right} onDismiss={this.onPanelClosed.bind(this)}></QualityCoordinatorForm>;
          case AllRoles.PTPAC_CHAIR.toUpperCase(): //PTPAChair Form
            return <PTPACCHAIRFORM isNewForm={false} invokedItem={this.state.selectedItem} context={this.props.context} position={PanelPosition.Right} onDismiss={this.onPanelClosed.bind(this)}></PTPACCHAIRFORM>;
          case AllRoles.PTPAC_REVIEWER.toUpperCase(): //PTPACReviewer Form
            return <PTPACReviwerForm isNewForm={false} invokedItem={this.state.selectedItem} context={this.props.context} position={PanelPosition.Right} onDismiss={this.onPanelClosed.bind(this)}></PTPACReviwerForm>;
          default:
            return null;
        }
      }
      else if (this.state.isNew) {
        return <IntakeFormPanel isNewForm={true} invokedItem={this.state.selectedItem} context={this.props.context} position={PanelPosition.Right} onDismiss={this.onPanelClosed.bind(this)}></IntakeFormPanel>;
      }
    }
  }
  private _copyAndSort<T>(items: T[], columnKey: string, isSortedDescending?: boolean): T[] {
    const key = columnKey as keyof T;
    return items.slice(0).sort((a: T, b: T) => ((isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1));
  }
  private _onRenderColumn(item: INominationListViewItem, index: number, column: IColumn) {
    const value =
      item && column && column.fieldName ? item[column.fieldName as keyof IDetailsListGroupedNominationItem] || '' : '';
    return <div data-is-focusable={true}>{value}</div>;
  }
  @autobind
  private _onColumnClick(ev: React.MouseEvent<HTMLElement>, column: IColumn): void {
    const { columns, pendingItems,completedItems } = this.state;
    const newColumns: IColumn[] = columns.slice();
    const currColumn: IColumn = newColumns.filter(currCol => column.key === currCol.key)[0];
    newColumns.forEach((newCol: IColumn) => {
      if (newCol === currColumn) {
        currColumn.isSortedDescending = !currColumn.isSortedDescending;
        currColumn.isSorted = true;
      } else {
        newCol.isSorted = false;
        newCol.isSortedDescending = true;
      }
    });
    const newPendingItems = this._copyAndSort(pendingItems, currColumn.fieldName!, currColumn.isSortedDescending);
    const neCompletedItems = this._copyAndSort(completedItems, currColumn.fieldName!, currColumn.isSortedDescending);
    this.setState({
      columns: newColumns,
      pendingItems: newPendingItems,
      completedItems: neCompletedItems,
    });
  }

  @autobind
  private _onItemInvoked(item: IDetailsListGroupedNominationItem): void {
    {
      const { masterItems } = this.state;
      if (masterItems) {
        const itemsSelected: INominationListViewItem[] = masterItems.filter(filterItem => filterItem.id === item.key);
        this.setState({
          isOpen: !this.state.isOpen,
          selectedItem: itemsSelected && itemsSelected.length > 0 ? itemsSelected[0] : null
        });
      }
    }
  }
  @autobind
  private onNewClick() {
    this.setState({
      isOpen: true,
      selectedItem: null,
      isNew: true

    });
  }
  private _getCommandBar(): JSX.Element {
    let _items: ICommandBarItemProps[] = [];
    let _farItems: ICommandBarItemProps[] = [];

    _items.push({
      key: 'newItem',
      text: 'Start a new nomination',
      iconProps: { iconName: 'Add' },
      split: true,
      ariaLabel: 'New',
      onClick: this.onNewClick,
    });

    return (
      <CommandBar
        items={_items}
        ariaLabel="Commands"
        farItems={_farItems}
        //className={styles.commandBarStyles}
        styles={commandBarStyles}
      />
    );
  }

  private _getEmptyCommandBar(): JSX.Element {
    let _items: ICommandBarItemProps[] = [];
    let _farItems: ICommandBarItemProps[] = [];

    return (
      <CommandBar
        items={[]}
        ariaLabel="Commands"
        farItems={[]}
        //className={styles.commandBarStyles}
        styles={commandBarStyles}
      />
    );
  }

  public render(): React.ReactElement<INominationProps> {
    //const formType = this.props.formType;
    const{pendingItems, completedItems ,columns, actor} = this.state;
    
    return (
      <div className={styles.nomination}>
        <ErrorBoundary>
          <div className={styles.container}>
            {actor !== null && actor.toUpperCase() == AllRoles.NOMINATOR.toUpperCase() ? this._getCommandBar() : this._getEmptyCommandBar()}      
            <>
              <DetailsList
                componentRef={this._root}
                items={pendingItems ? pendingItems: []}
                columns={this.state && this.state.columns}
                groupProps={{
                  showEmptyGroups: true,
                }}
                onRenderItemColumn={this._onRenderColumn}
                onItemInvoked={this._onItemInvoked}
                onRenderDetailsHeader={(headerProps, defaultRender) => {
                  return defaultRender({
                  ...headerProps,
                  styles: {
                    root: {
                    selectors: {
                      '.ms-DetailsHeader-cell': {
                        whiteSpace: 'normal',
                        textOverflow: 'clip',
                        lineHeight: 'normal',
                      },
                      '.ms-DetailsHeader-cellTitle': {
                        height: '100%',
                        alignItems: 'center',
                      },
                    },
                  }
                  }
                  });}}
              /> 
            </> 
          </div>
          {this.formPanel()}
          <div className={styles.QCContainer}>
            {/*   { !this.state.items.length && (
                <Stack horizontalAlign='center'>
                  <Text block>No item to display</Text>
                </Stack>
            )} */}        
          </div >
        </ErrorBoundary>
      </div >
    );
  }

  @autobind
  private onPanelClosed() {
    if (this.state && this.state.actor)
      this.setState({
        isOpen:false
      });
      this.initializeNominationsList(this.state.actor);
  }
}
