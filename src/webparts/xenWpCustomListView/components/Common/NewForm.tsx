import * as React from 'react';
import styles from '../XenWpCustomListView.module.scss';
import 'office-ui-fabric-core/dist/css/fabric.min.css';
import "@pnp/sp/fields";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/views";
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";

import { RichText } from "@pnp/spfx-controls-react/lib/RichText";
import "./CustomStyles/custom.css";
// import spService from './SPService/Service';
import spService from '../SPService/Service';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { ComboBox, DatePicker, DefaultButton, Dialog, DialogFooter, Dropdown, IComboBox, Icon, Label, PrimaryButton, Separator, TextField } from '@fluentui/react';
import { IPeoplePickerContext, PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
// import XenCustomPeoplePicker from './PresalesSPOControl';
// import basicDetails from './ColumnDetails.js'
// import { TextField } from '@fluentui/react';

interface IXenCreateFormProps {
    columnsDetails?: any;
    listName: any;
    siteUrl: any;
    context: WebPartContext;
    onCloseCreateForm: any;
}


interface IxenCreateFormState {
    Data: any;

    hideValidationError: boolean;
    hideSuccessDialog: boolean;
    hideFailueDialog: boolean;
    validationItems: any;
    basicDetailsColumn: any;
    projectDetailsColumn: any;
    otherDetailsColumn: any;
    customerColumnOption: any;
    attachments: any;
    filterUser: any;
    preSalesSPOUser: any

}

const basicDetails = [
    { title: "Title" },
    { title: "AccManager" },
    { title: "Segments" },
    { title: "CustomerName" },
    { title: "TPID" },
    { title: "Priority" },
    { title: "CustomerType" },
    { title: "DealType" },
    { title: "OpportunityType" },
    { title: "SalesPlay" },
    { title: "ContactPerson" },
    { title: "Designation" },
    { title: "Website" },
    // "Website"
    { title: "EmailId" },
    { title: "PhoneNo" },

    // "City"
    { title: "Region" },
    { title: "City" },
    { title: "LeadFrom" },
    { title: "MicrosoftSPOC" },

]



const projectDetailColumn = [
    { title: "ProjectType" },
    { title: "Status" },
    { title: "LeadDate" },
    { title: "ClosureDate" },
    { title: "NextFollowup" },
    { title: "ACR" },
    { title: "MRR" },
    { title: "MRR_x0024_" },
    { title: "AMM" },
    { title: "Value" },
    { title: "ManagedServices" },
    { title: "OTC" },
    { title: "ARR" },
    { title: "PCID" },
    { title: "RenewalDate" },
    { title: "PresalesSPOC" },
    { title: "DeliveryStatus" },

]
const otherDetailColumn = [
    { title: "Description" },
    { title: "NextAction" },
    { title: "DocStoreLink" },
    { title: "EngagementType" },
    { title: "Reminders" },
    { title: "ReminderStartDate" },
]

/* const presaleMappedUsers = [
    { user: [{ key: "vijai.b@xencia.com", text: "Vijai Bharrath", }], opptyType: "Managed Support" },
    { user: [{ key: "ikramul.haq@xencia.com", text: "Mohammed Ikram" }], opptyType: "App Innovation" },
    { user: [{ key: "aswath@xencia.com", text: "Aswath Narayanan" }], opptyType: "Analytics" },
    { user: [{ key: "aswath@xencia.com", text: "Aswath Narayanan" }], opptyType: "GenAI" },
    { user: [{ key: "lakshmanan.suresh@xencia.com", text: "Lakshmanan Suresh" }, { key: "kamlesh.giri@xencia.com", text: "Kamlesh Giri" }], opptyType: "IAAS" },
    { user: [{ key: "balaji.m@xencia.com", text: "Balaji M" }], opptyType: "Security" },
] */

    const textFieldFontSize = {
  
    root: {
      fontSize: 13, // placeholder font size
    },
    field:{
        fontSize:13
    }
  }


  const comboxFontSize =  {
          root: {
     
     
      fontSize:'13px'
    },
    input: {
     
      fontSize:'13px'
    },
    optionsContainerWrapper: {
     
      fontSize:'13px'
    },
      
      }



  const dropdownFontSize = {
    root: {
      fontSize: 13,
    },
    dropdown: {
      fontSize: 13,
    },
    title: {
      fontSize: 13,
    },
    callout: {
      fontSize: 13,
    },
    dropdownItem: {
      fontSize: 13,
    },
    dropdownItemSelected: {
      fontSize: 13,
    },
  }


  const datePickerFontSize = {
    root: {
      fontSize: 13,
    },
    textField: {
      fontSize: 13,
    },
    callout: {
      fontSize: 13,
    },
  }


export default class XenWpCreateForm extends React.Component<IXenCreateFormProps, IxenCreateFormState> {
    private _spService: spService;
    private _checkBoxItems: { [x: string]: string; }[] = [];
    private _peoplePickerContext: IPeoplePickerContext;
    private _fileRef: any;
    // protected presaleSPOC: any;
    constructor(props: IXenCreateFormProps) {
        super(props);
        this.state = {
            Data: {},
            hideValidationError: true,
            hideSuccessDialog: true,
            hideFailueDialog: true,
            validationItems: [],
            customerColumnOption: [],
          /*   basicDetailsColumn: this._arrangeDetailsColumn(basicDetails) || [],
            projectDetailsColumn: this._arrangeDetailsColumn(projectDetailColumn) || [],//otherDetailsColumn
            otherDetailsColumn: this._arrangeDetailsColumn(otherDetailColumn) || [],//otherDetailsColumn */
              basicDetailsColumn: [],
            projectDetailsColumn:  [],//otherDetailsColumn
            otherDetailsColumn:  [],//otherDetailsColumn
            attachments: [],
            filterUser: [],
            preSalesSPOUser: ""
        }

        this._fileRef = React.createRef<HTMLInputElement>();
        this._peoplePickerContext = {
            absoluteUrl: this.props.context.pageContext.web.absoluteUrl,
            msGraphClientFactory: this.props.context.msGraphClientFactory,
            spHttpClient: this.props.context.spHttpClient
        };

        const siteInfo = this.props.siteUrl || '';
        this._spService = new spService(this.props.context, siteInfo);
        this._getCustmerColumnOption("");
    }

    /*   private  _intinalPreSalesDropDwnOption=async ()=>{
          const options= await this._spService.getPreSaleSPoOptions("")||[];
          this.setState({filterUser:options})
      } */

    private _arrangeDetailsColumn = (ColumnDetails: any) => {
        // Merge matching items
        const merged = ColumnDetails
            .map((detail: any) => {
                const match = this.props.columnsDetails.find((field: any) => field.Title.trim() === detail.title.trim());
                return match ? { ...match } : null;
            })
            .filter((item: any) => item !== null);
        console.log(merged,"merged")

        return merged; 
    }
    // get checkbox valuesstartsWith
    public handleChxBoxChange = (eve: React.ChangeEvent<HTMLInputElement> | React.ChangeEvent<HTMLSelectElement>, index: number): void => {
        const { name, value } = eve.target;
        const fltrCh = this._checkBoxItems.map(x => x[name])
        if (!fltrCh.includes(value)) {
            this._checkBoxItems.push({
                [name]: value
            });


        } else {

            const itemToRemoveIndex = this._checkBoxItems.findIndex(function (item) {
                return item[name] === value;
            });

            // proceed to remove an item only if it exists.
            if (itemToRemoveIndex !== -1) {
                this._checkBoxItems.splice(itemToRemoveIndex, 1);
            }
        }
        const fltrCheckItems = this._checkBoxItems.map(x => {
            if (typeof x[name] !== "undefined") {
                return x[name]
            }
        })
        const CheckTemp: string[] = []
        fltrCheckItems.filter(ele => {
            if (typeof ele !== "undefined") {
                CheckTemp.push(ele);
            }
        });
        this.setState(prevState => ({
            Data: {
                ...prevState.Data, [name]: CheckTemp.sort(),

            }
        }));
    }

    componentDidMount(): void {
        this.setState({
            basicDetailsColumn:this._arrangeDetailsColumn(basicDetails) || [],
            projectDetailsColumn: this._arrangeDetailsColumn(projectDetailColumn) || [],//otherDetailsColumn
            otherDetailsColumn:this._arrangeDetailsColumn(otherDetailColumn) || [],//otherDetailsColumn
        })
  
       
         
        
    }

    private _getPeoplePickerItemsSingleUser = (nm: string, items: any[]) => {

        console.log(nm, items)
        if (items.length > 0) {
            this.setState(prevState => ({ Data: { ...prevState.Data, [nm]: items[0]?.id } }));
        } else {

            this.setState(prevState => ({ Data: { ...prevState.Data, [nm]: null } }));

        }
    }


 /*    private _getPeoplePickerItems(nm: string, items: any[]): void {
        const apprIds: number[] = [];
        const apprEmails: string[] = []
        const item = items;
        for (let i = 0; i < item.length; i++) {
            //    id ..........
            if (!apprIds.includes(item[i].id)) {
                apprIds.push(item[i].id);
            }
            else {
                const index = apprIds.indexOf(item[i].id);
                if (index > -1) {
                    apprIds.splice(index, 1);
                }
            }

            // emails..................................
            if (!apprEmails.includes(item[i].secondaryText)) {
                apprEmails.push(item[i].secondaryText);
            }
            else {
                const index = apprEmails.indexOf(item[i].secondaryText);
                if (index > -1) {
                    apprEmails.splice(index, 1);
                }
            }
        }
        this.setState(prevState => ({ Data: { ...prevState.Data, [nm]: apprIds } }));
    } */

    /* Test Field OnChange */
    private _handleTextFieldChnage = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: any) => {
        const currentElementId = event.currentTarget.id || "";
        if (currentElementId) {
            this.setState(prevState => ({ Data: { ...prevState.Data, [currentElementId]: newValue } }));
        }
    }

    /*  */
    // get test ,radio, dropdown values eventHandlerBoolean
    public eventHandler = (eve: React.ChangeEvent<HTMLInputElement> | React.ChangeEvent<HTMLSelectElement>, index: number): void => {
        const { name, value } = eve.target;
        this.setState(prevState => ({ Data: { ...prevState.Data, [name]: value } }));
    }

    // get textarea values
    public handleTextareaChange = (event: React.ChangeEvent<HTMLTextAreaElement>, ind?: number): void => {
        const { name, value } = event.target;
        this.setState(prevState => ({ Data: { ...prevState.Data, [name]: value } }));
    }

    public _getCustmerColumnOption = async (value: any) => {
        //  this. _intinalPreSalesDropDwnOption();
        const respose = await this._spService.getCustomerDrpDwnOption("");
        if (respose) {
            this.setState({ customerColumnOption: respose })
        }
    }

    /* DropDown Change  */
    private _OnChangeDrpDown = (event: React.FormEvent<HTMLDivElement>, option?: any, index?: number) => {
        const name = event.currentTarget.id;
        const internalName = name.split("-")[0];
        if (internalName) {
            this.setState(prevState => ({ Data: { ...prevState.Data, [internalName]: option?.key } }),
                () => {
                    if (internalName && internalName === "OpportunityType") {
                        this.setState(prevState => ({ Data: { ...prevState.Data, ["PresalesSPOCId"]: null } }));
                        this._filterPresalesUserDynamically(option?.key);
                    }
                }
            );
        }
    }

    /* DropDown Change  */
    private _getPreSalesDropChange = (event: React.FormEvent<HTMLDivElement>, option?: any, index?: number) => {
        this.setState(prevState => ({ Data: { ...prevState.Data, ["PresalesSPOCId"]: option?.key } }));
    }

    /* Get presales options */
    private _filterPresalesUserDynamically = async (value: string) => {
        const findFilterUsers = await this._spService.getPreSaleSPoOptions(value);
        if (findFilterUsers) {
            this.setState({ filterUser: findFilterUsers, preSalesSPOUser: findFilterUsers[0]?.key });
            this.setState(prevState => ({ Data: { ...prevState.Data, ["PresalesSPOCId"]: findFilterUsers[0]?.key } }));
        } else {
            this.setState({ filterUser: [], preSalesSPOUser: null });
        }
    }


    ///* Attachembts  */
    private addAttacment = async (): Promise<void> => {
        const fileInfo: { name: string; content: File; index: number; fileUrl: string; ServerRelativeUrl: string; isExists: boolean; Modified: string; isSelected: boolean; }[] = [];
        const fileInput: any = document.getElementById('Docfiles') as HTMLInputElement;
        const fileCount = fileInput.files.length;
        for (let i = 0; i < fileCount; i++) {
            // const file = fileInput["files"][i];
            const file = fileInput.files[i];
            const filesId = Math.floor((Math.random() * 1000000000) + 1);
            const reader = new FileReader();
            reader.onload = ((file) => {
                return (e) => {
                    //Push the converted file into array
                    e.preventDefault();
                    const isObjectExists = this.state.attachments.map((obJ: { name: string; }) => obJ.name);
                    if (!isObjectExists.includes(file.name)) {
                        fileInfo.push({
                            "name": file.name,
                            "content": file,
                            "index": filesId,
                            "fileUrl": "",
                            "ServerRelativeUrl": "",
                            "isExists": false,
                            "Modified": new Date().toISOString(),
                            "isSelected": false
                        });
                    }
                    this.setState({ attachments: [...this.state.attachments, ...fileInfo] });
                    console.log(fileInfo, this.state.attachments)
                    // this.fileInfos.push(fileInfo);
                };
            })(file);
            reader.readAsArrayBuffer(file);
        }
    }

    // Remove Attachemnts
    public onRemoveAttachments = (file: any): void => {
        const { attachments } = this.state;
        const fltrArry = attachments.filter((obj: { index: any; }) => obj.index !== file.index);
        this.setState({ attachments: fltrArry });
    }

    private _onSubmit = async () => {
        const items = this.state.Data;
        const response = await this._spService.addCustomerdata(this.props.listName, items, this.state.attachments)
        if (response) {
            this.setState({ hideSuccessDialog: !this.state.hideSuccessDialog })
        } else {
            this.setState({ hideFailueDialog: !this.state.hideFailueDialog })
        }
    }

    private _validation = () => {
        const columnInfo = [...this.state.basicDetailsColumn, ...this.state.projectDetailsColumn, ...this.state.otherDetailsColumn]
        const object = this.state.Data;
        let errorObj: string[] = [];

        const filterRequiredFields = columnInfo?.filter((col: { Required: boolean }) => col.Required);

        if (Object.keys(object).length === 0) {
            // If no data, mark all required fields as missing
            filterRequiredFields.forEach((col: { text: string }) => {
                errorObj.push(col.text);
            });
        } else {
            filterRequiredFields.forEach((col: { internalName: string, text: string }) => {
                const value = object[col.internalName];

                // If key doesn't exist or value is empty/null/undefined
                if (value === undefined || value === null || value === "") {
                    errorObj.push(col.text);
                }
            });
        }
        if (errorObj?.length > 0) {
            this.setState({ hideValidationError: false, validationItems: errorObj });
        } else {
            this._onSubmit();
        }

        console.log("Missing Required Fields:", errorObj);
    };

    private onCloseFormDialog = () => {
        this.setState({
            hideFailueDialog: true,
            hideSuccessDialog: true,
            hideValidationError: true
        })
        this.props.onCloseCreateForm();
    }

    private _OnChangeComboBox = (event: React.FormEvent<IComboBox>, option?: any, index?: number, value?: string) => {
        this.setState(prevState => ({ Data: { ...prevState.Data, [`CustomerNameId`]: option?.key } }));
    }


    private _handleDatePicker = (date: any, field: string) => {
        this.setState(prevState => ({ Data: { ...prevState.Data, [field]: date } }));
    }

    private _onRichTextChange = (test: any, internalname: string) => {
        this.setState(prevState => ({ Data: { ...prevState.Data, [internalname]: test } }));
        return test;
    }

  

    private _configureColumnRender = (columnInfo: any) => {
        // console.log(columnInfo)
        // debugger;
        const columnsDetails = columnInfo.map((_x: any, index: number) => {
            if (_x.dataType === "Text") {
                return (
                    <div
                    //  className={`ms-Grid-col ms-sm12 ms-md6 ms-xl4 ${styles._projectDetailsEachContainer}`}
                      className={`${styles._projectDetailsEachContainer}`}
                     >
                        <div className={styles.fieldEditor}>
                            <div className={styles._customLabelContainer}>
                                <Icon iconName="TextField" />

                                <Label className='label' htmlFor={_x.internalName}>{_x.text}{_x.Required ? <span className={styles._customRequiredlabel}>*</span> : null}</Label>

                            </div>
                           

                            <TextField 
                             styles={textFieldFontSize}
                            iconProps={{ iconName: "Text" }} id={_x.internalName} onChange={this._handleTextFieldChnage} value={_x.text === "Customer Name" ? this.state.Data[`Customer Name`] : this.state.Data[_x.internalName] || ""} placeholder={`Enter ${_x.text}`}/>
                            

                             {/* <br/>
                            <span style={{color:'green',fontStyle:'oblique'}}><strong>1</strong></span> */}
                              {_x.text === 'Project Name' &&<span className={`${styles._spanDanger}`}>You Can't leave this blank</span>}
                             {/* {_x.text === 'Project Name' &&<br/>} */}
                             {_x.text === 'Project Name' &&<span className={`${styles._spanDescription}`}>Kindly use following template:</span>}
                            {/* {_x.text === 'Project Name' &&<br/>} */}
                            {_x.text === 'Project Name' &&<span className={`${styles._spanDescription}`}>{`<<Customer First Name>> <<Opportunity Name>> <<MMYY>>`}</span>}
                        </div>
                    </div>
                )
            }
        /*     else if (_x.dataType === "Choice") {
                return (
                    <div className="ms-Grid-col ms-sm12 ms-md6 ms-xl4">
                        <div className={styles.fieldEditor}>
                            <div className={styles._customLabelContainer}>
                                <Icon iconName='CheckboxComposite' />
                                <Label className='label' htmlFor={_x.internalName}>{_x.text}{_x.Required ? <span className={styles._customRequiredlabel}>*</span> : null}</Label>
                            </div>
                            <div className={styles.radioCheckbox}>
                                {_x.option.map((ele: string | number | readonly string[]) => {
                                    return (
                                        <div key={index} className={styles.radiocontainer}>
                                            <input type="radio" name={_x.internalName} value={ele}
                                                onChange={(event) => this.eventHandler(event, index)}

                                                checked={this.state.Data[_x.internalName] === ele ? true : false}
                                            />
                                            <div className={styles.radiocheckmark}>
                                                <span className={styles.radioinsidecircle} />
                                            </div>
                                            {ele}
                                        </div>
                                    )
                                })}

                            </div>
                        </div>
                    </div>
                )
            } */
            //"Lookup"
            else if (_x.dataType === "Lookup" && _x.Title === "CustomerName") {
                // debugger
                return (
                    <div
                    //  className={`ms-Grid-col ms-sm12 ms-md6 ms-xl4 ${styles._projectDetailsEachContainer}`}
                      className={`${styles._projectDetailsEachContainer}`}
                    >
                        <div className={styles.fieldEditor}>
                            <div className={styles._customLabelContainer}>
                                <Icon iconName="Split" />

                                <Label className='label' htmlFor={_x.internalName}>{_x.text}{_x.Required ? <span className={styles._customRequiredlabel}>*</span> : null}</Label>

                            </div>

                            <ComboBox id={_x.internalName} options={this.state.customerColumnOption}
                                onChange={this._OnChangeComboBox}
                                selectedKey={this.state.Data[_x.internalName] || ""}  placeholder={`Enter ${_x.text}`}
                                
                                 styles={comboxFontSize}
/>
                                
                                {/* <span>2</span> */}
                        </div>
                    </div>
                )
            }
        /*     else if (_x.dataType === "Boolean") {
                return (
                    <div>
                        <div className={styles.fieldEditor}>
                            <div className={styles._customLabelContainer}>
                                <Icon iconName='CheckboxComposite' />
                                <Label className='label' htmlFor={_x.internalName}>{_x.text}{_x.Required ? <span className={styles._customRequiredlabel}>*</span> : null}</Label>
                            </div>
                            <div className='radioCheckbox'>
                                {_x.option.map((ele: any) => {
                                    return (
                                        <div key={index} className={styles.radiocontainer}>
                                            <input type="radio" name={_x.internalName} value={ele}
                                                onChange={(event) => this.eventHandler(event, index)}
                                                checked={this.state.Data[_x.internalName] === ele ? true : false}
                                            />
                                            <div className={styles.radiocheckmark}>
                                                <span className={styles.radioinsidecircle} />
                                            </div>
                                            {ele === "true" ? "Yes" : "No"}
                                        </div>
                                    )
                                })}
                            </div>
                        </div>
                    </div>
                )
            } */
          /*   else if (_x.dataType === "MultiChoice") {
                return (<div className="ms-Grid-col ms-sm12 ms-md6 ms-xl4">

                    <div className={styles.fieldEditor}>
                        <div className={styles._customLabelContainer}>
                            <Icon iconName='CheckboxComposite' />
                            <Label className='label' htmlFor={_x.internalName}>{_x.text}{_x.Required ? <span className={styles._customRequiredlabel}>*</span> : null}</Label>
                        </div>
                        <div className='radioCheckbox'>
                            {_x.option.map((ele: string) => {
                                return (
                                    // <label htmlFor={value.internalName} key={index}>{ele}
                                    <div key={index} className={styles._CheckBoxcontainer}>

                                        <input type="checkbox" name={_x.internalName} value={ele}
                                            onChange={(event) => this.handleChxBoxChange(event, index)}
                                            checked={this._checkBoxItems.some(obj => obj[_x.internalName] === ele)}

                                        />
                                        <span className={styles.checkmark} />
                                        {ele}

                                    </div>
                                    // </label>
                                )
                            })}

                        </div>
                    </div>
                </div>)
            } */
            else if (_x.dataType === "Dropdown") {
                return (<div
                //  className={`ms-Grid-col ms-sm12 ms-md6 ms-xl4 ${styles._projectDetailsEachContainer}`}
                  className={`${styles._projectDetailsEachContainer}`}
                  >
                    {/* <div className={styles.fieldEditor}> */}
                        <div className={styles._customLabelContainer}>
                            <Icon iconName='CheckboxComposite' />
                            <Label className='label' htmlFor={_x.internalName}>{_x.text}{_x.Required ? <span className={styles._customRequiredlabel}>*</span> : null}</Label>
                        </div>

                        <Dropdown id={_x.internalName} options={_x.option}
                        placeholder={`Select ${_x.text}`}

                            onChange={this._OnChangeDrpDown}
                            selectedKey={this.state.Data[_x.internalName] || ""}
                           styles={dropdownFontSize}
                            
                           
                        />
                        {/* <span>3</span> */}
                        {/* <br/> */}
                        {_x.text ==="Lead From" && <span className={`${styles._spanDescription}`}>Provide Partner name in Description.</span>}
                    {/* </div> */}
                </div>)
            }

            else if (_x.dataType === "DateTime") {
                return (<div
                //  className={`ms-Grid-col ms-sm12 ms-md6 ms-xl4 ${styles._projectDetailsEachContainer}`}
                  className={`${styles._projectDetailsEachContainer}`}
                  >
                    {/* <div className={styles.fieldEditor}> */}
                        <div className={styles._customLabelContainer}>
                            <Icon iconName='Calendar' />
                            <Label className='label' htmlFor={_x.internalName}>{_x.text}{_x.Required ? <span className={styles._customRequiredlabel}>*</span> : null}</Label>
                        </div>
                        <DatePicker
                            id={_x.internalName}
                            minDate={_x.Title === "LeadDate" ? new Date() : undefined}
                            value={this.state.Data[_x.internalName] ? new Date(this.state.Data[_x.internalName]) : undefined}
                            onSelectDate={(date) => this._handleDatePicker(date, _x.internalName)}
                            placeholder={`Select ${_x.text}`} 
                        // onChange={this._handleTextFieldChnage} 
                        styles={datePickerFontSize}

                        />
                        {/* <span>4</span> */}
                    {/* </div> */}
                </div>)
            }

            else if (_x.dataType === "Number") {
                return (<div
                //  className={`ms-Grid-col ms-sm12 ms-md6 ms-xl4 ${styles._projectDetailsEachContainer}`}
                  className={`${styles._projectDetailsEachContainer}`}
                  >
                    {/* <div className={styles.fieldEditor}> */}
                        <div className={styles._customLabelContainer}>
                            <Icon iconName='NumberField' />
                            <Label className='label' htmlFor={_x.internalName}>{_x.text}{_x.Required ? <span className={styles._customRequiredlabel}>*</span> : null}</Label>
                        </div>

                        
                        
                        <TextField styles={textFieldFontSize} type="number" id={_x.internalName} onChange={this._handleTextFieldChnage} value={this.state.Data[_x.internalName] || ""} 

                        
                        placeholder={`Enter ${_x.text}`} />
                        {/* <span>5</span> */}
                    {/* </div> */}
                </div>)
            }
         else if (_x.dataType === "Currency") {
                return (<div
                //  className={`ms-Grid-col ms-sm12 ms-md6 ms-xl4 ${styles._projectDetailsEachContainer}`}
                 
                  className={`${styles._projectDetailsEachContainer}`}
                  >
                    <div className={styles.fieldEditor}>
                        <div className={styles._customLabelContainer}>
                            <Icon iconName='NumberField' />
                            <Label className='label' htmlFor={_x.internalName}>{_x.text}{_x.Required ? <span className={styles._customRequiredlabel}>*</span> : null}</Label>
                        </div>
                        <div>
                            <TextField type="number" id={_x.internalName} onChange={this._handleTextFieldChnage} value={this.state.Data[_x.internalName] || ""}
                             styles={textFieldFontSize}
                            placeholder={`Enter ${_x.text}`} />
                            {/* <span>6</span> */}
                        </div>

                    </div>
                </div>)
            } 
         /*    else if (_x.dataType === "UserMulti") {
                return (<div className="ms-Grid-col ms-sm12 ms-md6 ms-xl4">
                    <div className={styles.fieldEditor}>
                        <div className={styles._customLabelContainer}>
                            <Icon iconName='Contact' />
                            <Label className='label' htmlFor={_x.internalName}>{_x.text}{_x.Required ? <span className={styles._customRequiredlabel}>*</span> : null}</Label>
                        </div>
                        <PeoplePicker
                            context={this._peoplePickerContext}
                            personSelectionLimit={3}
                            groupName={""} // Leave this blank in case you want to filter from all users
                            showtooltip={true}

                            disabled={false}
                            onChange={this._getPeoplePickerItems.bind(this, _x.internalName)}
                            showHiddenInUI={false}
                            ensureUser={true}
                            principalTypes={[PrincipalType.User]}
                            resolveDelay={1000} />

                    </div>
                </div>)
            } */
            else if (_x.dataType === "User") {
                return (<div
                //  className={`ms-Grid-col ms-sm12 ms-md6 ms-xl4 ${styles._projectDetailsEachContainer}`}
                  className={`${styles._projectDetailsEachContainer}`}
                  >
                    {/* <div className={styles.fieldEditor}> */}

                        <div className={styles._customLabelContainer}>
                            <Icon iconName='Contact' />
                            <Label className='label' htmlFor={_x.internalName}>{_x.text}{_x.Required ? <span className={styles._customRequiredlabel}>*</span> : null}</Label>
                        </div>
                        {_x.Title === "PresalesSPOC" ? <div><Dropdown
                        styles={dropdownFontSize}
                            options={this.state.filterUser}
                            placeholder={`Enter ${_x.text}`} 
                            defaultSelectedKey={this.state.preSalesSPOUser}
                            selectedKey={this.state.Data[_x.internalName] || ""}
                            onChange={this._getPreSalesDropChange}
                        />
                        {/* <span>7</span> */}
                        </div>
                            :
                            <PeoplePicker
                                context={this._peoplePickerContext}
                                placeholder={`Enter ${_x.text}`} 
                                personSelectionLimit={1}
                                groupName={""} // Leave this blank in case you want to filter from all users
                                showtooltip={true}
                                disabled={false}
                                onChange={this._getPeoplePickerItemsSingleUser.bind(this, _x.internalName)}
                                showHiddenInUI={false}
                                ensureUser={true}
                                principalTypes={[PrincipalType.User]}
                                resolveDelay={1000} />}

                    {/* </div> */}
                </div>
                )
            }

            else if (_x.dataType === "Note") {
                return (
                  <div
  className={`richContainer ${styles._projectDetailsEachContainer}`}
>
  <div className={styles._customLabelContainer}>
    <Icon iconName="AlignLeft" />
    <Label className="label" htmlFor={_x.internalName}>
      {_x.text}
      {_x.Required && (
        <span className={styles._customRequiredlabel}>*</span>
      )}
    </Label>
  </div>

  {_x.internalName === "DocStoreLink" ? (
    <div className={styles.richTextFieldWrapper}>
         <RichText
        onChange={(text) => this._onRichTextChange(text, _x.internalName)}
        value={this.state.Data[_x.internalName] || ""}
      />


    </div>

     
   
  ) : (
    <TextField
      multiline
      rows={2}
      id={_x.internalName}
      onChange={this._handleTextFieldChnage}
      value={this.state.Data[_x.internalName] || ""}
      placeholder={`Enter ${_x.text}`}
      styles={textFieldFontSize}
    />
  )}

  {_x.text === "Description" && (
    <span className={styles._spanDescription}>
      Any description related to the project
    </span>
  )}
</div>


                )
            }
            /*   else if (_x.dataType === "Note") {
                  return (
                      <div className="ms-Grid-col ms-sm12 ms-md6 ms-xl4">
                          <div className={styles.fieldEditor} key={index}>
                              <div className={styles._customLabelContainer}>
                                  <Icon iconName='AlignLeft' />
                                  <Label className='label' htmlFor={_x.internalName}>{_x.text}{_x.Required ? <span className={styles._customRequiredlabel}>*</span> : null}</Label>
                              </div>
                              {_x.internalName === "DocStoreLink" && (<RichText onChange={(text) => this._onRichTextChange(text, _x.internalName)} value={this.state.Data[_x.internalName] || ""} /> )}
  
  
                          </div>
                      </div>
  
                  )
              }  */
          // RichText
            else {
                return (
                    <div
                    //  className={`ms-Grid-col ms-sm12 ms-md6 ms-xl4 ${styles._projectDetailsEachContainer}`}
                      className={`${styles._projectDetailsEachContainer}`}
                      >
                        <div className={styles.fieldEditor}>
                            <div className={styles._customLabelContainer}>
                                <Icon iconName="TextField" />

                                <Label className='label' htmlFor={_x.internalName}>{_x.text}{_x.Required ? <span className={styles._customRequiredlabel}>*</span> : null}</Label>

                            </div>

                            <TextField 
                             styles={textFieldFontSize} id={_x.internalName} onChange={this._handleTextFieldChnage} value={this.state.Data[_x.internalName] || ""}  placeholder={`Enter ${_x.text}`} />
                            {/* <span>9</span> */}
                        </div>
                    </div>
                )
            }
 
        })
        return columnsDetails;
    }

       private _configureProjectDetaail = (columnInfo: any) => {
        // console.log(columnInfo)
        // debugger;
        const columnsDetails = columnInfo.map((_x: any, index: number) => {
            if (_x.dataType === "Text") {
                return (
                    
                    <div 
                    // className={`ms-Grid-col ms-sm12 ms-md6 ms-xl4 ${styles._projectDetailsEachContainer}`}
                     className={`${styles._projectDetailsEachContainer}`}
                     key={index}  >
                        {/* <div className={styles.fieldEditor}> */}
                            <div className={styles._customLabelContainer}>
                                <Icon iconName="TextField" />

                                <Label className='label' htmlFor={_x.internalName}>{_x.text}{_x.Required ? <span className={styles._customRequiredlabel}>*</span> : null}</Label>

                            </div>

                            <TextField  styles={textFieldFontSize}   id={_x.internalName} onChange={this._handleTextFieldChnage} value={this.state.Data[_x.internalName] || ""} placeholder={`Enter ${_x.text}`}  />
                            {/* <span>1</span> */}
                            

                        {/* </div> */}
                    </div>
                )
            }
        /*     else if (_x.dataType === "Choice") {
                return (
                    <div className="ms-Grid-col ms-sm12 ms-md6 ms-xl4">
                        <div className={styles.fieldEditor}>
                            <div className={styles._customLabelContainer}>
                                <Icon iconName='CheckboxComposite' />
                                <Label className='label' htmlFor={_x.internalName}>{_x.text}{_x.Required ? <span className={styles._customRequiredlabel}>*</span> : null}</Label>
                            </div>
                            <div className={styles.radioCheckbox}>
                                {_x.option.map((ele: string | number | readonly string[]) => {
                                    return (
                                        <div key={index} className={styles.radiocontainer}>
                                            <input type="radio" name={_x.internalName} value={ele}
                                                onChange={(event) => this.eventHandler(event, index)}

                                                checked={this.state.Data[_x.internalName] === ele ? true : false}
                                            />
                                            <div className={styles.radiocheckmark}>
                                                <span className={styles.radioinsidecircle} />
                                            </div>
                                            {ele}
                                        </div>
                                    )
                                })}

                            </div>
                        </div>
                    </div>
                )
            } */
            //"Lookup"
            else if (_x.dataType === "Lookup" && _x.Title === "CustomerName") {
                // debugger
                return (
                    <div
                    //  className={`ms-Grid-col ms-sm12 ms-md6 ms-xl4 ${styles._projectDetailsEachContainer}`}
                      className={`${styles._projectDetailsEachContainer}`}
                      key={index} >
                        <div className={styles.fieldEditor}>
                            <div className={styles._customLabelContainer}>
                                <Icon iconName="Split" />

                                <Label className='label' htmlFor={_x.internalName}>{_x.text}{_x.Required ? <span className={styles._customRequiredlabel}>*</span> : null}</Label>

                            </div>

                            <ComboBox id={_x.internalName} options={this.state.customerColumnOption}  styles={comboxFontSize}
                                onChange={this._OnChangeComboBox}
                                selectedKey={this.state.Data[_x.internalName] || ""} placeholder={`Enter ${_x.text}`} />
                        </div>
                        {/* <span>2</span> */}
                    </div>
                )
            }
        /*     else if (_x.dataType === "Boolean") {
                return (
                    <div>
                        <div className={styles.fieldEditor}>
                            <div className={styles._customLabelContainer}>
                                <Icon iconName='CheckboxComposite' />
                                <Label className='label' htmlFor={_x.internalName}>{_x.text}{_x.Required ? <span className={styles._customRequiredlabel}>*</span> : null}</Label>
                            </div>
                            <div className='radioCheckbox'>
                                {_x.option.map((ele: any) => {
                                    return (
                                        <div key={index} className={styles.radiocontainer}>
                                            <input type="radio" name={_x.internalName} value={ele}
                                                onChange={(event) => this.eventHandler(event, index)}
                                                checked={this.state.Data[_x.internalName] === ele ? true : false}
                                            />
                                            <div className={styles.radiocheckmark}>
                                                <span className={styles.radioinsidecircle} />
                                            </div>
                                            {ele === "true" ? "Yes" : "No"}
                                        </div>
                                    )
                                })}
                            </div>
                        </div>
                    </div>
                )
            } */
          /*   else if (_x.dataType === "MultiChoice") {
                return (<div className="ms-Grid-col ms-sm12 ms-md6 ms-xl4">

                    <div className={styles.fieldEditor}>
                        <div className={styles._customLabelContainer}>
                            <Icon iconName='CheckboxComposite' />
                            <Label className='label' htmlFor={_x.internalName}>{_x.text}{_x.Required ? <span className={styles._customRequiredlabel}>*</span> : null}</Label>
                        </div>
                        <div className='radioCheckbox'>
                            {_x.option.map((ele: string) => {
                                return (
                                    // <label htmlFor={value.internalName} key={index}>{ele}
                                    <div key={index} className={styles._CheckBoxcontainer}>

                                        <input type="checkbox" name={_x.internalName} value={ele}
                                            onChange={(event) => this.handleChxBoxChange(event, index)}
                                            checked={this._checkBoxItems.some(obj => obj[_x.internalName] === ele)}

                                        />
                                        <span className={styles.checkmark} />
                                        {ele}

                                    </div>
                                    // </label>
                                )
                            })}

                        </div>
                    </div>
                </div>)
            } */
            else if (_x.dataType === "Dropdown") {
                return (<div
                //  className={`ms-Grid-col ms-sm12 ms-md6 ms-xl4 ${styles._projectDetailsEachContainer}`} 
                  className={`${styles._projectDetailsEachContainer}`}
                 key={index} >
                    {/* <div className={styles.fieldEditor}> */}
                        <div className={styles._customLabelContainer}>
                            <Icon iconName='CheckboxComposite' />
                            <Label className='label' htmlFor={_x.internalName}>{_x.text}{_x.Required ? <span className={styles._customRequiredlabel}>*</span> : null}</Label>
                        </div>

                        <Dropdown id={_x.internalName} options={_x.option}
                        placeholder={`Select ${_x.text}`} 
styles={dropdownFontSize}
                            onChange={this._OnChangeDrpDown}
                            selectedKey={this.state.Data[_x.internalName] || ""}
                        />
                        {/* <span>3</span> */}
                    {/* </div> */}
                </div>)
            }

            else if (_x.dataType === "DateTime") {
                return (<div 
                // className={`ms-Grid-col ms-sm12 ms-md6 ms-xl4 ${styles._projectDetailsEachContainer}`} 
                 className={`${styles._projectDetailsEachContainer}`} key={index} 
                >
                    {/* <div className={styles.fieldEditor}> */}
                        <div className={styles._customLabelContainer}>
                            <Icon iconName='Calendar' />
                            <Label className='label' htmlFor={_x.internalName}>{_x.text}{_x.Required ? <span className={styles._customRequiredlabel}>*</span> : null}</Label>
                        </div>
                        <DatePicker
                         styles={datePickerFontSize}
                            id={_x.internalName}
                            placeholder={`Enter ${_x.text}`} 
                            minDate={_x.Title === "LeadDate" ? new Date() : undefined}
                            value={this.state.Data[_x.internalName] ? new Date(this.state.Data[_x.internalName]) : undefined}
                            onSelectDate={(date) => this._handleDatePicker(date, _x.internalName)}
                        // onChange={this._handleTextFieldChnage} 

                        />
                        {/* <span>4</span> */}
                        {/* <br/> */}
                         {_x.text === 'Renewal Date' &&<span className={`${styles._spanDescription}`}>Support or renwal date</span>}
                    {/* </div> */}
                </div>)
            }

            else if (_x.dataType === "Number") {
                return (<div
                //  className={`ms-Grid-col ms-sm12 ms-md6 ms-xl4 ${styles._projectDetailsEachContainer}`}
                  className={`${styles._projectDetailsEachContainer}`}
                   key={index} >
                    <div className={styles.fieldEditor}>
                        <div className={styles._customLabelContainer}>
                            <Icon iconName='NumberField' />
                            <Label className='label' htmlFor={_x.internalName}>{_x.text}{_x.Required ? <span className={styles._customRequiredlabel}>*</span> : null}</Label>
                        </div>

                        <TextField  styles={textFieldFontSize}  type="number" id={_x.internalName} onChange={this._handleTextFieldChnage} value={this.state.Data[_x.internalName] || ""} placeholder={`Enter ${_x.text}`} />
                        {/* <span>5</span> */}
                        {/* <br/> */}
                        {_x.text === 'Azure Consumption Revenue (Annual) ($)' &&<span className={`${styles._spanDescription}`}>In Dollar</span>}
                        {}
                    </div>
                </div>)
            }
         else if (_x.dataType === "Currency") {
                return (<div
                //  className={`ms-Grid-col ms-sm12 ms-md6 ms-xl4 ${styles._projectDetailsEachContainer}`} 
                  className={`${styles._projectDetailsEachContainer}`}
                   key={index} >
                    <div className={styles.fieldEditor}>
                        <div className={styles._customLabelContainer}>
                            <Icon iconName='NumberField' />
                            <Label className='label' htmlFor={_x.internalName}>{_x.text}{_x.Required ? <span className={styles._customRequiredlabel}>*</span> : null}</Label>
                        </div>
                        <div>
                            <TextField  styles={textFieldFontSize}  type="number" id={_x.internalName} onChange={this._handleTextFieldChnage} value={this.state.Data[_x.internalName] || ""} placeholder={`Enter ${_x.text}`} />
                            {/* <span>6</span> */}
                            {/* <br/> */}
                            {_x.text === 'Azure Consumption Revenue (Annual) ($)' &&<span className={`${styles._spanDescription}`}>In Dollar</span>}
                            {_x.text === 'MRR (₹)' &&<span className={`${styles._spanDescription}`}>In INR</span>}
                            {_x.text === 'MRR ($)' &&<span className={`${styles._spanDescription}`}>In Dollar</span>}
                            {_x.text === 'AMM ($)' &&<span className={`${styles._spanDescription}`}>In Dollar</span>}
                            {_x.text === 'Services Value (₹)' &&<span className={`${styles._spanDescription}`}>In INR</span>}
                            {_x.text === 'Managed Services (₹)' &&<span className={`${styles._spanDescription}`}>In INR</span>}
                            {_x.text === 'OTC (₹)' &&<span className={`${styles._spanDescription}`}>In INR</span>}
                            {_x.text === 'Annual Recurring Revenue (₹)' &&<span className={`${styles._spanDescription}`}>In INR</span>}
                           
                        </div>

                    </div>
                </div>)
            } 
         /*    else if (_x.dataType === "UserMulti") {
                return (<div className="ms-Grid-col ms-sm12 ms-md6 ms-xl4">
                    <div className={styles.fieldEditor}>
                        <div className={styles._customLabelContainer}>
                            <Icon iconName='Contact' />
                            <Label className='label' htmlFor={_x.internalName}>{_x.text}{_x.Required ? <span className={styles._customRequiredlabel}>*</span> : null}</Label>
                        </div>
                        <PeoplePicker
                            context={this._peoplePickerContext}
                            personSelectionLimit={3}
                            groupName={""} // Leave this blank in case you want to filter from all users
                            showtooltip={true}

                            disabled={false}
                            onChange={this._getPeoplePickerItems.bind(this, _x.internalName)}
                            showHiddenInUI={false}
                            ensureUser={true}
                            principalTypes={[PrincipalType.User]}
                            resolveDelay={1000} />

                    </div>
                </div>)
            } */
            else if (_x.dataType === "User") {
                return (<div
                //  className={`ms-Grid-col ms-sm12 ms-md6 ms-xl4 ${styles._projectDetailsEachContainer}`}
                  className={`${styles._projectDetailsEachContainer}`}
                   key={index} >
                    {/* <div className={styles.fieldEditor}> */}

                        <div className={styles._customLabelContainer}>
                            <Icon iconName='Contact' />
                            <Label className='label' htmlFor={_x.internalName}>{_x.text}{_x.Required ? <span className={styles._customRequiredlabel}>*</span> : null}</Label>
                        </div>
                        {_x.Title === "PresalesSPOC" ? <div><Dropdown
                        placeholder={`Enter ${_x.text}`} 
                            options={this.state.filterUser}
                            defaultSelectedKey={this.state.preSalesSPOUser}
                            selectedKey={this.state.Data[_x.internalName] || ""}
                            onChange={this._getPreSalesDropChange}
                            styles={dropdownFontSize}
                        />
                        {/* <span>7</span> */}
                        </div>
                            :
                            <PeoplePicker
                            placeholder={`Enter ${_x.text}`} 
                                context={this._peoplePickerContext}
                                personSelectionLimit={1}
                                groupName={""} // Leave this blank in case you want to filter from all users
                                showtooltip={true}
                                disabled={false}
                                onChange={this._getPeoplePickerItemsSingleUser.bind(this, _x.internalName)}
                                showHiddenInUI={false}
                                ensureUser={true}
                                principalTypes={[PrincipalType.User]}
                                resolveDelay={1000} />}

                    {/* </div> */}
                </div>
                )
            }

            else if (_x.dataType === "Note") {
                return (
                    <div
                    //  className={`ms-Grid-col ms-sm12 ms-md6 ms-xl4 ${styles._projectDetailsEachContainer}`} 
                      className={`${styles._projectDetailsEachContainer}`}
                       key={index} >
                        {/* <div className={styles.fieldEditor} key={index}> */}
                            <div className={styles._customLabelContainer}>
                                <Icon iconName='AlignLeft' />
                                <Label className='label' htmlFor={_x.internalName}>{_x.text}{_x.Required ? <span className={styles._customRequiredlabel}>*</span> : null}</Label>
                            </div>
                            {_x.internalName === "DocStoreLink" ?
                             <div className={styles.richTextFieldWrapper}>
                            <RichText onChange={(text) => this._onRichTextChange(text, _x.internalName)} value={this.state.Data[_x.internalName] || ""} /> </div> :

                                <TextField  styles={textFieldFontSize}  multiline={true} rows={2} id={_x.internalName} onChange={this._handleTextFieldChnage} value={this.state.Data[_x.internalName] || ""} placeholder={`Enter ${_x.text}`} />}
                                    {/* <span>8</span> */}
                        {/* </div> */}
                    </div>

                )
            }
            /*   else if (_x.dataType === "Note") {
                  return (
                      <div className="ms-Grid-col ms-sm12 ms-md6 ms-xl4">
                          <div className={styles.fieldEditor} key={index}>
                              <div className={styles._customLabelContainer}>
                                  <Icon iconName='AlignLeft' />
                                  <Label className='label' htmlFor={_x.internalName}>{_x.text}{_x.Required ? <span className={styles._customRequiredlabel}>*</span> : null}</Label>
                              </div>
                              {_x.internalName === "DocStoreLink" && (<RichText onChange={(text) => this._onRichTextChange(text, _x.internalName)} value={this.state.Data[_x.internalName] || ""} /> )}
  
  
                          </div>
                      </div>
  
                  )
              }  */
          // RichText
            else {
                return (
                    <div 
                    // className={`ms-Grid-col ms-sm12 ms-md6 ms-xl4 ${styles._projectDetailsEachContainer}`} 
                     className={`${styles._projectDetailsEachContainer}`}
                      key={index} >
                        <div className={styles.fieldEditor}>
                            <div className={styles._customLabelContainer}>
                                <Icon iconName="TextField" />

                                <Label className='label' htmlFor={_x.internalName}>{_x.text}{_x.Required ? <span className={styles._customRequiredlabel}>*</span> : null}</Label>

                            </div>

                            <TextField  styles={textFieldFontSize}  id={_x.internalName} onChange={this._handleTextFieldChnage} value={this.state.Data[_x.internalName] || ""} placeholder={`Enter ${_x.text}`} />
                            {/* <span>9</span> */}
                        </div>
                    </div>
                )
            }
 
        })
        return columnsDetails;
    }

    /*     private presalesSPO = () => {
            return (<PeoplePicker
                context={this._peoplePickerContext}
                personSelectionLimit={1}
                suggestionsLimit={1}
                groupName={""} // Leave this blank in case you want to filter from all users
                showtooltip={true}
                resultFilter={(result) => this._filterUserInPeoplePicker(result)}
                disabled={false}
                onChange={this._getPeoplePickerItemsSingleUser.bind(this,"PresalesSPOCId")}
                showHiddenInUI={false}
                ensureUser={true}
                ref={c => (this.presaleSPOC = c)}
                principalTypes={[PrincipalType.User]}
                resolveDelay={1000} />)
        } */

    /*   private _filterUserInPeoplePicker = (result: any) => {
          const { filterUser } = this.state;
          if (filterUser.length > 0) {
              return result.filter((p: any) => filterUser.includes(p["secondaryText"].toLowerCase()));
          } else {
              return result
          }
      } */

    public render(): React.ReactElement<IXenCreateFormProps> {
        // console.log(this.state)
        const { validationItems, basicDetailsColumn, projectDetailsColumn, otherDetailsColumn } = this.state;
        return (
            <div style={{ padding: "20px",overscrollBehavior: "contain" ,
 overflow: "hidden" }} className="ms-Grid">

                <div className="ms-Grid-row">
                    <h2 className={styles.marginBottom}>Oppty number -</h2>
                </div>
                <Separator />
                <div style={{ height: "51vh", overflow: "auto", padding: "20px" }}>
                    <div className="ms-Grid-row">
                        <h3 className={styles.marginTop}>Basic Details</h3>
                    </div>
                    
                        <div
                        //  className="ms-Grid-row" 
                        className={styles.flexContainer}
                        >
                        {this._configureColumnRender(basicDetailsColumn)}
                    </div>

                    

                    
                    <Separator />

                    <div className="ms-Grid-row">
                        <h3>Project Details</h3>
                    </div>
                    <div 
                    // className="ms-Grid-row"
                     className={styles.flexContainer}
                    // style={{display:'flex',flexWrap:'wrap',gap:'16px'}}
                    >
                        {/* <div className='ms-Grid-col'> */}
                        {this._configureProjectDetaail(projectDetailsColumn)}
                        {/* </div> */}
                    </div>
                    <Separator />

                    <div className="ms-Grid-row">
                        <h3>Other Details</h3>
                    </div>
                    <div
                    //  className="ms-Grid-row"
                      className={styles.flexContainer}
                     >
                        {this._configureColumnRender(otherDetailsColumn)}
                    </div>
                    <Separator />
                    <div className="ms-Grid-row">
                        <div className="ms-Grid-col ms-sm12 ms-md6 ms-xl4">
                            <div className={styles.fieldEditor}>
                                <div className={styles._customLabelContainer}>
                                    <Icon iconName='Attach' />
                                    <Label className='label' htmlFor="attchments">Attachments</Label>
                                </div>
                                <div className={styles._attachementContiner}>
                                    {this.state.attachments.length > 0 && (this.state.attachments.map((_x: any, index: number) => {
                                        return (<div className={styles._attchemntContainerLabel}><span>{_x.name}</span><Icon iconName='Cancel' onClick={() => this.onRemoveAttachments(_x)} /></div>)
                                    }))}
                                </div>
                                <DefaultButton type='button' data-is-focusable={false}    text='Add Attachments' 
                                onClick={(e:any) => {
                                    e.preventDefault(); // Prevent default behavior
        e.stopPropagation(); // Stop event bubbling
                                    if (this._fileRef) {
                                        const fileclick = this._fileRef!
                                        fileclick.current.click()
                                    }
                                }
                                } 
                                />

                            </div>

                        </div>
                    </div>
                </div>

                <Separator />
                <div className={styles._customBtnContainer}>
                    <span>
                        <PrimaryButton text="Submit" onClick={this._validation} />
                    </span><span>
                        <DefaultButton text="Cancel" onClick={this.onCloseFormDialog} />
                    </span>

                </div>
                <Dialog
                    onDismiss={() => this.setState({ hideValidationError: !this.state.hideValidationError })}
                    hidden={this.state.hideValidationError}

                    dialogContentProps={{
                        title: "Alert",

                    }}
                >
                    <div>
                        <p>Please filed the all required fields</p>
                        {
                            validationItems?.map((_x: any) => {
                                return <li>{_x}</li>
                            })
                        }
                    </div>
                    <DialogFooter>
                        <PrimaryButton onClick={() => this.setState({ hideValidationError: !this.state.hideValidationError })} text="Ok" />
                    </DialogFooter>
                </Dialog>

                <Dialog
                    onDismiss={() => this.setState({ hideSuccessDialog: !this.state.hideSuccessDialog })}
                    hidden={this.state.hideSuccessDialog}

                    dialogContentProps={{
                        title: "Alert",

                    }}
                >
                    <div>
                        <p>Request has been submitted sucessfully.</p>

                    </div>
                    <DialogFooter>
                        <PrimaryButton onClick={this.onCloseFormDialog} text="Ok" />
                    </DialogFooter>
                </Dialog>
                <Dialog
                    onDismiss={() => this.setState({ hideFailueDialog: !this.state.hideFailueDialog })}
                    hidden={this.state.hideFailueDialog}

                    dialogContentProps={{
                        title: "Alert",

                    }}
                >
                    <div>
                        <p>Something went wrong. Please try again.</p>

                    </div>
                    <DialogFooter>
                        <PrimaryButton onClick={() => this.setState({ hideFailueDialog: !this.state.hideFailueDialog })} text="Ok" />
                    </DialogFooter>
                </Dialog>

                <input type='file' onChange={this.addAttacment} hidden={true} ref={this._fileRef} id={"Docfiles"} />


            </div >
        );
    }
}