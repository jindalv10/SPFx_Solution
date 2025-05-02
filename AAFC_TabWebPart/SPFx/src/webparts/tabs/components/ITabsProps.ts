export interface ITabsProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  sectionClass: string;
  webPartClass: string;
  tabData: any[];
  children?: any;
  change?: 'click' | 'hover';
  displayMode: any;
  jqueryDomElement: any;
}
