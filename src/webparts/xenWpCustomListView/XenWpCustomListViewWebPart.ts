import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  // PropertyPaneTextField,
  PropertyPaneDropdown,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import * as strings from 'XenWpCustomListViewWebPartStrings';
import XenWpCustomListView from './components/XenWpCustomListView';
import { IXenWpCustomListViewProps } from './components/IXenWpCustomListViewProps';
import { PropertyFieldSitePicker } from '@pnp/spfx-property-controls/lib/PropertyFieldSitePicker';
import { spfi, SPFx } from "@pnp/sp";
import spService from '../xenWpCustomListView/components/SPService/Service'
import "@pnp/sp/webs";
import "@pnp/sp/lists";

export interface IXenWpCustomListViewWebPartProps {
  description: string;
  site:any;
  list:any;
  listOption:any;
  isSortingEnable:boolean;
  isSearchEnable:boolean;
  filterColumnName:string;
  customColumnNameOption:any
}

export default class XenWpCustomListViewWebPart extends BaseClientSideWebPart<IXenWpCustomListViewWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _sp:any;
  private _listDrpDwnOptions:any=[];
  // private _spSerive:any;
  public render(): void {
    const element: React.ReactElement<IXenWpCustomListViewProps> = React.createElement(
      XenWpCustomListView,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        context:this.context,
        list:this.properties.list,
        site:this.properties.site,
        listOption:this.properties.listOption,
        isSortingEnable:this.properties.isSortingEnable,
        isSearchEnable:this.properties.isSearchEnable,
        filterColumnName:this.properties.filterColumnName,
  customColumnNameOption:this.properties.customColumnNameOption,

      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    if(this.properties.site){
      // this._spSerive =new spService(this.context,this.properties.site[0]?.url||[]);
      this.properties.listOption= await this.getlistItem(this.properties.site)||[]
    }
    if(this.properties.site && this.properties.list){
      this.properties.customColumnNameOption =await this.getListColumns( this.properties.list||"")
    }

  
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }

  private getlistItem=async (customSiteUrl?:any)=>{
    if(this.properties.site && customSiteUrl){
      // const siteUrl= this.properties.site[0]?.url
    this._sp=spfi(customSiteUrl[0]?.url ||"https://xencia1.sharepoint.com/sites/XenciaSalesTracker/").using(SPFx(this.context));
   
    const lists = await this._sp.web.lists.filter(`BaseTemplate eq ${100}`)();
    console.log(lists)
    if(lists){
      lists.map((_x: any)=>{
        this._listDrpDwnOptions.push({key:_x.Title,text:_x.Title});
      });
      return this._listDrpDwnOptions;
    }

    }
  }



  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected onPropertyPaneFieldSiteChanged=async (propertyPath: string, oldValue: any, newValue: any)=>{
    debugger;
if(newValue !== oldValue){
  this._listDrpDwnOptions=[];
  this.properties.list=null;
  await this.getlistItem(newValue)
}
  }

  private async getListColumns(listTitle: string): Promise<any[]> {
    if (this.properties.site && listTitle) {
      const service = new spService(this.context, this.properties.site[0]?.url);
      const fields = await service.getfieldInfo(listTitle) || [];
      // Transform fields into dropdown options
      return fields.map(field => ({
        key: field.key,
        text: field.text
      }));
    }
    return [];
  }

  // ... [keep other existing methods unchanged] ...

  protected onPropertyPaneFieldChanged = async (propertyPath: string, oldValue: any, newValue: any): Promise<void> => {
    if (propertyPath === 'site') {
      this.properties.list = "";
      this.properties.listOption = await this.getlistItem(newValue);
      this.properties.customColumnNameOption = []; // Reset columns when site changes
      this.render();
      this.context.propertyPane.refresh();
    }
    else if (propertyPath === 'list') {
      if (this.properties.site && newValue) {
        this.properties.customColumnNameOption = await this.getListColumns(newValue);
        this.properties.filterColumnName = ""; // Reset selected column
        this.context.propertyPane.refresh();
      }
    }
    else {
      super.onPropertyPaneFieldChanged(propertyPath, newValue, oldValue);
    }
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    this.getlistItem(this.properties.site)
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyFieldSitePicker('site', {
                  label: 'Select sites',
                  initialSites: this.properties.site,
                  context: this.context,
                  deferredValidationTime: 500,
                  multiSelect: false,
                  onPropertyChange: this.onPropertyPaneFieldSiteChanged.bind(this),
                  properties: this.properties,
                  key: 'sitesFieldId'
                }),
              PropertyPaneDropdown('list',{
                label:"select a list",
                options:this._listDrpDwnOptions,
                selectedKey:this.properties.list
              }),
              PropertyPaneDropdown('filterColumnName',{
                label:"select a column",
                options:this.properties.customColumnNameOption,
                selectedKey:this.properties.filterColumnName
              }),
              PropertyPaneToggle("isSortingEnable",{
                label:"Enable Sorting"
              }),
              PropertyPaneToggle("isSearchEnable",{
                label:"Enable Search"
              })
              ]
            }
          ]
        }
      ]
    };
  }

  
}
