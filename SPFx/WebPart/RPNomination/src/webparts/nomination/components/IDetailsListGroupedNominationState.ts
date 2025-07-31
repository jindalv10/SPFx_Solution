import { IColumn, IGroup } from '@fluentui/react';
import { INominationListViewItem } from 'pd-nomination-library';
import { IDetailsListGroupedNominationItem } from './IDetailsListGroupedNominationItem';

export interface IDetailsListGroupedNominationState {
  pendingItems:IDetailsListGroupedNominationItem [];
  completedItems:IDetailsListGroupedNominationItem[];
  masterItems: INominationListViewItem[];
  isOpen?: boolean;
  selectedItem: INominationListViewItem;
  actor: string;
  columns: IColumn[];
  isNew: boolean;

}