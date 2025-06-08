import * as React from 'react';
import styles from './XenWpCustomListView.module.scss';
import 'office-ui-fabric-core/dist/css/fabric.min.css';
import type { IXenWpCustomListViewProps } from './IXenWpCustomListViewProps';
import { spfi, SPFx } from "@pnp/sp";
import { IRenderListDataParameters } from "@pnp/sp/lists";
import "@pnp/sp/fields";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/views";
import spService from './SPService/Service';
import { ListView, SelectionMode } from "@pnp/spfx-controls-react/lib/ListView";
import {
  CommandBar,
  // Dialog,
  ICommandBarItemProps,

  Modal,
  Persona,
  PersonaSize,
  // ScrollablePane,
  SearchBox,
  Spinner,
  SpinnerSize
} from '@fluentui/react';
import { MSGraphClientV3 } from "@microsoft/sp-http";
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import XenWpCreateForm from './Common/NewForm';
import XenWpEditForm from './Common/EditForm';

interface IXenWpCustomListViewState {
  allItems: any[];
  rawdata: any[];
  viewColumn: any[];
  isLoading: boolean;
  currentUserEmail: string;
  managerEmail: string;
  managerDirectReports: string[];
  columnInfo: any;
  hideCreateForm: boolean;
  hideEditForm: boolean;
  selectionDetails: any;
  selectedcount: number;

}




export default class XenWpCustomListView extends React.Component<IXenWpCustomListViewProps, IXenWpCustomListViewState> {
  private _sp;
  private _spService: spService;
  private _siteUrl: any;
  private _selection: Selection;

  constructor(props: IXenWpCustomListViewProps) {
    super(props);
    this.state = {
      allItems: [],
      viewColumn: [],
      isLoading: true,
      currentUserEmail: '',
      managerEmail: '',
      managerDirectReports: [],
      rawdata: [],
      columnInfo: [],
      hideCreateForm: false,
      selectionDetails: [],
      selectedcount: 0,
      hideEditForm: false

    };
    const siteInfo = this.props.site || [];
    this._siteUrl = siteInfo[0]?.url || "";
    this._spService = new spService(this.props.context, siteInfo[0]?.url);
    this._sp = spfi(siteInfo[0]?.url).using(SPFx(this.props.context));
    console.log(this._selection)
    console.log(this.props)
    // Initialize the component by fetching user data and items
    this._initializeComponent();
  }

  private _initializeComponent = async (): Promise<void> => {
    try {
      // First, get the current user's manager and their direct reports
      await this._getUserManagerDetails();
      // Then get the list items
      await this._GetAllItems();
    } catch (error) {
      console.error("Error initializing component:", error);
      this.setState({ isLoading: false });
    }
  }

  private _getUserManagerDetails = async (): Promise<void> => {
    try {
      // Get an instance of the MSGraphClient
      const graphClient: MSGraphClientV3 = await this.props.context.msGraphClientFactory.getClient("3");
      // Get the current user's email
      const currentUserEmail = this.props.context.pageContext.user.email;
      const managerResponse = await graphClient
        .api('/me')
        .select('mail,displayName,id')
        .get() as MicrosoftGraph.User;

      if (!managerResponse?.mail) {
        console.log("No manager found for current user");
        return;
      }

      console.log("User's manager:", managerResponse);
      // const managerEmail = managerResponse.mail;

      // Get all direct reports of the manager (including the current user and their peers)
      const directReportsResponse = await graphClient
        .api(`/users/${managerResponse.id}/directReports`)
        .select('mail,displayName')
        .get() as { value: MicrosoftGraph.User[] };

      const managerDirectReports: string[] = [currentUserEmail];

      // Add all direct reports of the manager
      if (directReportsResponse?.value?.length > 0) {
        directReportsResponse.value.forEach((user) => {
          if (user.mail) {
            managerDirectReports.push(user.mail);
          }
        });
      } /* else {
        // If no direct reports, just add current user
        managerDirectReports.push(currentUserEmail);
      } */

      this.setState({
        managerDirectReports: managerDirectReports
      });

    } catch (error) {
      console.error("Error in user manager details:", error);
      // Fallback to just current user if any error occurs
      const currentUserEmail = this.props.context.pageContext.user.email;
      this.setState({
        currentUserEmail,
        managerDirectReports: [currentUserEmail]
      });
    }
  }

  private _GetAllItems = async (): Promise<void> => {
    try {
      /* get all items */
      const view = await this._sp.web.lists.getByTitle(this.props.list).views.getByTitle("All Items")();
      await this._getViewId(view.Id);

      const renderListDataParams: IRenderListDataParameters = {
        ViewXml: `<View><Query>${view.ViewQuery}</Query></View>`,
      };

      const r = await this._sp.web.lists.getByTitle(this.props.list).renderListDataAsStream(renderListDataParams);
      const allItems = r.Row || [];

      // Filter items based on AccountManager
      const allHaveEmail = allItems.every(manager => manager?.hasOwnProperty(this.props.filterColumnName || ""));
      let filterItems = []
      if (allHaveEmail) {
        filterItems = allItems?.filter((item: any) => this._FilterAccountsManger(item[this.props.filterColumnName], allHaveEmail));
      }
      else {
        filterItems = allItems;
      }
      this.setState({
        allItems: filterItems,
        rawdata: filterItems,
        isLoading: false
      });
    } catch (error) {
      console.error("Error getting list items:", error);
      this.setState({ isLoading: false });
    }
  }

  private _FilterAccountsManger = (users: any, allHaveEmail: boolean): boolean => {
    let isValid = true;
    const reportingDirectors = this.state.managerDirectReports || []
    // If users is undefined, null, or empty, return 
    if (allHaveEmail) {
      if (typeof users === "object") {
        isValid = users.some((user: any) => reportingDirectors.includes(user.email));
      } else {
        isValid = false
      }
    }
    return isValid
  }

  private _getViewId = async (id: any): Promise<void> => {
    try {
      const fieldsDeatils = await this._spService.getfieldDetails(this.props.list);
      const allViewFields: any[] = [];

      if (fieldsDeatils) {
        fieldsDeatils.forEach((_x: any) => {
          if (_x.dataType === "UserMulti" || _x.dataType === "User") {
            allViewFields.push({
              name: _x.Title,
              displayName: _x.text,
              minWidth: 200,
              maxWidth: 300,
              isResizable: true,
              render: (item?: any) => {
                const people = [];
                let i = 0;
                while (item[`${_x.Title}.${i}.title`]) {
                  people.push({
                    title: item[`${_x.Title}.${i}.title`],
                    email: item[`${_x.Title}.${i}.email`],
                  });
                  i++;
                }
                return (people.map((_person: any) =>
                  <Persona
                    key={_person.email}
                    size={PersonaSize.size24}
                    showInitialsUntilImageLoads
                    imageShouldStartVisible
                    imageUrl={`/_layouts/15/userphoto.aspx?username=${_person.email}&size=M`}
                  >
                    <span style={{ fontSize: "12px" }}>{_person.title}</span>
                  </Persona>
                ));
              }
            });
          } else if (_x.isRichText === "true" ||_x.isRichText === true) {
            allViewFields.push({
              name: _x.internalName,
              displayName: _x.text,
              minWidth: 500,
              maxWidth: 500,
              isResizable: true,
              sorting: false,

              render: (item?: any, index?: number, column?: any) => {
                return <div dangerouslySetInnerHTML={{ __html: item[_x.internalName] }}></div>
              }
            })

          }
          else {
            allViewFields.push({
              name: _x.internalName,
              displayName: _x.text,
              minWidth: 200,
              maxWidth: 300,
              isResizable: true,
              sorting: this.props.isSortingEnable,

            });
          }
        });

        this.setState({ viewColumn: allViewFields, columnInfo: fieldsDeatils });
      }
    } catch (error) {
      console.error("Error getting view fields:", error);
    }
  }


  private _sortitems = (items: any[], columnName: any, descending: boolean) => {
    const columnInfo = this.state.columnInfo;
    const columnType = columnInfo?.find((_x: { internalName: any; }) => _x.internalName === columnName)?.dataType
    let sortedItems = []
    if (this.props.isSortingEnable) {
      // debugger;

      sortedItems = [...items].sort((a, b) => {
        let aValue = a[columnName]
        let bValue = b[columnName];
        // Handle null/undefined values (place them at the end)
        if (aValue === null && bValue === null) return 0;
        if (aValue === null) return 1;  // nulls last
        if (bValue === null) return -1; // nulls last

        // Handle date comparison
        if (columnType === "DateTime" || aValue instanceof Date) {
          // Convert to Date objects if they aren't already
          const aDate = aValue instanceof Date ? aValue : new Date(aValue);
          const bDate = bValue instanceof Date ? bValue : new Date(bValue);

          // Use getTime() for comparison
          aValue = aDate.getTime();
          bValue = bDate.getTime();
        }
        if (aValue === bValue) return 0;

        if (!descending) {
          return aValue > bValue ? 1 : -1;
        } else {
          return aValue < bValue ? 1 : -1;
        }
      });

    } else {
      sortedItems = [...items]
    }
    return sortedItems
  }


  private onSearch = (text: string): void => {

    this.setState({
      allItems: text ? this.state.rawdata.filter((item: any) => Object.values(item as Record<string, any>).some(value =>
        (value || '').toString().toLowerCase().indexOf(text.toLowerCase()) > -1
      ))

        : this.state.rawdata,
    });
    // }
  }

  private onClear = () => {
    this.setState({
      allItems: this.state.rawdata
    })
  }

  private _onDismissCreateFormDialog = () => {
    this.setState({ hideCreateForm: !this.state.hideCreateForm })
  }
  private _onDismissEditFormDialog = () => {
    this.setState({ hideEditForm: !this.state.hideEditForm })
  }
  //

  public _getSelection = async (item: any) => {


    const fileterUserFileds = this.state.columnInfo.filter((_x: any) => _x.dataType === "UserMulti" || _x.dataType === "User");
    const userfields = fileterUserFileds.map((_x: { Title: any; }) => _x.Title);
    const updatedArray = userfields.map((item: any) => `${item}/EMail`);
    const selectFileds = updatedArray.join(",");
    const expandfields = userfields.join(",");
    const currentObj = await this._spService.getItemById(this.props.list, Number(item[0].ID), selectFileds, expandfields);
    const contentObj = currentObj?.items;
    const filecotent = currentObj?.files;
    fileterUserFileds.map((_x: any) => {
      contentObj[_x.Title] = [contentObj[_x.Title]?.EMail];

    });
    contentObj["_Files"] = filecotent;
    // contentObj["CustomerName"] = ""
    this.setState({ selectionDetails: contentObj, selectedcount: item?.length || 0 });

  }

  // private getCustomerName=(item:any)=>{


  // }

  public render(): React.ReactElement<IXenWpCustomListViewProps> {
    console.log(this.state)
    const {
      hasTeamsContext,
    } = this.props;
    const _items: ICommandBarItemProps[] = [
      {
        key: 'newItem',
        text: 'New Request',
        iconProps: { iconName: "Add" },
        onClick: () => this.setState({ hideCreateForm: true })

      },
      {
        key: 'editItem',
        text: 'Edit',
        disabled: this.state.selectedcount !== 1,
        iconProps: { iconName: "Edit" },
        // onClick: () => this.setState({ hideEditForm: this.state.selectedcount === 1 ? false : true })
        onClick: () => this.setState({ hideEditForm: this.state.selectedcount === 1 ? true : false })
      }
    ];

    const { isLoading, allItems, viewColumn } = this.state;

    return (
      <section className={`${styles.xenWpCustomListView} ${hasTeamsContext ? styles.teams : ''}`}>
        <CommandBar items={_items} />
        {this.props.isSearchEnable && (<div className={styles._CustomSearchContoiiner}>
          <span>Search </span>
          <SearchBox onSearch={this.onSearch} onClear={this.onClear} title='Search' />
        </div>)}
        <div className={styles.custom_list}>
          {isLoading ? (
            <Spinner size={SpinnerSize.large} label="Loading items..." />
          ) : (
            <ListView
              items={allItems}
              viewFields={viewColumn}
              compact={true}
              selectionMode={SelectionMode.single}
              selection={this._getSelection}
              stickyHeader={false}
              sortItems={this._sortitems}
            />
          )}
        </div>

        <Modal
          isOpen={this.state.hideCreateForm}
          onDismiss={() => this.setState({ hideCreateForm: !this.state.hideCreateForm })}
          isBlocking={false}
           styles={{
        scrollableContent: {
            overflow: "hidden" // Prevent default scrolling behavior
        }
    }}
          
        >

          <XenWpCreateForm
            columnsDetails={this.state.columnInfo}
            listName={this.props.list || ""}
            context={this.props.context}
            siteUrl={this._siteUrl}
            onCloseCreateForm={this._onDismissCreateFormDialog}
          />

        </Modal>
        {/*   <Dialog
          hidden={this.state.hideCreateForm}
          maxWidth={1000}
          minWidth={300}

          onDismiss={() => this.setState({ hideCreateForm: !this.state.hideCreateForm })}
        >
          <XenWpCreateForm
            columnsDetails={this.state.columnInfo}
            listName={this.props.list || ""}
            context={this.props.context}
            siteUrl={this._siteUrl}
            onCloseCreateForm={this._onDismissCreateFormDialog}
          />
        </Dialog> */}

        <Modal
          isOpen={this.state.hideEditForm}
          onDismiss={() => this.setState({ hideEditForm: !this.state.hideEditForm })}
          isBlocking={false}
        >
          <XenWpEditForm
            data={this.state.selectionDetails}
            columnsDetails={this.state.columnInfo}
            listName={this.props.list || ""}
            context={this.props.context}
            siteUrl={this._siteUrl}
            onCloseCreateForm={this._onDismissEditFormDialog}
          />

        </Modal>

        {/*    <Dialog
          hidden={this.state.hideEditForm}
          maxWidth={1000}
          minWidth={300}

          onDismiss={() => this.setState({ hideEditForm: !this.state.hideEditForm })}
        >

          <XenWpEditForm
            data={this.state.selectionDetails}
            columnsDetails={this.state.columnInfo}
            listName={this.props.list || ""}
            context={this.props.context}
            siteUrl={this._siteUrl}
            onCloseCreateForm={this._onDismissEditFormDialog}
          />
        </Dialog> */}

      </section>
    );
  }
}