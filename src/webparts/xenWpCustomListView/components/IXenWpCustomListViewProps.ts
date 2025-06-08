import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IXenWpCustomListViewProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context:WebPartContext;
  site:any;
  list:any;
  listOption:any;
  isSortingEnable:boolean;
  isSearchEnable:boolean;
  filterColumnName:string;
  customColumnNameOption:any;

}
