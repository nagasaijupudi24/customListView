import { WebPartContext } from "@microsoft/sp-webpart-base";
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/fields";
import "@pnp/sp/views";
import "@pnp/sp/items";
import "@pnp/sp/attachments";
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";



import { IItem } from "@pnp/sp/items";
export interface IKeyItems{ 
    key: string;
     text: string;
      dataType: string,
       option?: any[]; 
       internalName: string;
        DefaultValue?: any;
         FillInChoice?: boolean;
          Required: boolean,
           Title: string 
        isRichText?:boolean}
export interface IApproCheck {
    status: string;
    ActionTaken: string
}

export default class spService {
    private _sp;
    constructor(private context: WebPartContext, siteUrl: any) {
        this._sp = spfi(siteUrl).using(SPFx(this.context))
    }

    /* Get dropdown options  */
    private _getDropdownOption = (options: any) => {
        let renderOption: any = [];
        if (options && typeof options === "object" && options?.length > 0) {
            options?.map((_x: any) => {
                renderOption.push({
                    key: _x,
                    text: _x
                })
            })
        }

        return renderOption;



    }


    public getfieldDetails = async (listName: string): Promise<IKeyItems[]> => {
        const temp: IKeyItems[] = []
        await this._sp.web.lists.getByTitle(listName).fields.filter("Hidden eq false and ReadOnlyField eq false")().then(field => {
            field.filter(async (value: {
                DefaultValue?: any;
                Required: boolean;

                InternalName: string; Title: string; TypeDisplayName: string; Choices?: string[], TypeAsString: string, SchemaXml?: string; FillInChoice?: boolean
            }) => {
                if (!(
                    value.InternalName === "Attachments" ||
                    value.InternalName === "ContentType")) {
                    // newData.dataType === "UserMulti" || newData.dataType === "User")
                    // console.log("first"); StartProcessing
                    if (value.TypeAsString === "UserMulti" || value.TypeAsString === "User") {
                        temp.push({
                            key: value.Title,
                            text: value.Title,
                            dataType: value.TypeAsString,
                            internalName: value.InternalName + "Id",
                            Required: value.Required,
                            Title: value.InternalName,
                             isRichText:false,
                            // FillInChoice:value.FillInChoice
                        })
                    } else
                        if (value.TypeAsString === "Choice" && value?.SchemaXml?.match(/Dropdown/)) {
                            temp.push({
                                key: value.Title,
                                text: value.Title,
                                dataType: "Dropdown",
                                option: this._getDropdownOption(value.Choices),
                                internalName: value.InternalName,
                                Required: value.Required,
                                Title: value.InternalName,
                                 isRichText:false,
                                // FillInChoice:value.FillInChoice
                            })
                        }
                        else if (value.TypeAsString === "Choice" || value.TypeAsString === "MultiChoice") {
                            temp.push({
                                key: value.Title,
                                text: value.Title,
                                dataType: value.TypeAsString,
                                option: value.Choices,
                                internalName: value.InternalName,
                                DefaultValue: value.DefaultValue,
                                FillInChoice: value.FillInChoice,
                                Required: value.Required,
                                Title: value.InternalName,
                                 isRichText:false,
                            })
                        }
                        else if (value.TypeAsString === "Boolean") {
                            temp.push({
                                key: value.Title,
                                text: value.Title,
                                dataType: value.TypeAsString,
                                option: ["true", "false"],
                                internalName: value.InternalName,
                                Required: value.Required,
                                Title: value.InternalName,
                                 isRichText:false,
                            })

                        }
                        else if (value.TypeAsString === "Note") {
                            const xmlString = value?.SchemaXml || "";

                            // Parse the XML string
                            const parser = new DOMParser();
                            const xmlDoc = parser.parseFromString(xmlString, "text/xml");

                            // Get the <Field> element
                            const fieldElement = xmlDoc.querySelector("Field");

                            // Check the RichText attribute
                            const isRichText = fieldElement?.getAttribute("RichText") === "TRUE";
                  
                                temp.push({
                                    key: value.Title,
                                    text: value.Title,
                                    isRichText:isRichText,
                                    dataType: value.TypeAsString,
                                    internalName: value.InternalName,
                                    Required: value.Required,
                                    Title: value.InternalName
                                })
                        }
                        else if (value.TypeAsString === "Lookup" && value.InternalName === "CustomerName") {
                            temp.push({
                                key: value.Title,
                                text: value.Title,
                                dataType: value.TypeAsString,
                                option: [],
                                 isRichText:false,
                                internalName: value.InternalName + "Id",
                                Required: value.Required,
                                Title: value.InternalName
                            })

                        }
                        else {
                            temp.push({
                                key: value.Title,
                                text: value.Title,
                                 isRichText:false,
                                dataType: value.TypeAsString,
                                internalName: value.InternalName,
                                Required: value.Required,
                                Title: value.InternalName
                            })

                        }
                }
            })
        });
        return temp
    }

    public getfieldInfo = async (listName: string): Promise<{ key?: string; text?: string; }[]> => {
        const temp: { key: string; text: string; }[] = []
        await this._sp.web.lists.getByTitle(listName).fields.filter("Hidden eq false and ReadOnlyField eq false")().then(field => {
            field.filter((value: {
                DefaultValue?: any;
                InternalName: string; Title: string; TypeDisplayName: string; Choices?: string[], TypeAsString: string, SchemaXml?: string; FillInChoice?: boolean
            }) => {
                if (!(
                    value.InternalName === "Attachments" ||
                    value.InternalName === "ContentType")) {
                    if (value.TypeAsString === "UserMulti" || value.TypeAsString === "User") {
                        temp.push({
                            key: value.InternalName,
                            text: value.Title,
                        })
                    }
                }
            })
        });
        return temp;
    }

    public addCustomerdata = async (listName: string, obj: any, attachments: any): Promise<boolean> => {
        try {
            const item: any = await this._sp.web.lists.getByTitle(listName).items.add(obj);
            if (attachments?.length > 0) {

                const itemById: IItem = await this._sp.web.lists.getByTitle(listName).items.getById(item.ID);
                attachments?.map(async (_x: any) => {
                    await itemById.attachmentFiles.add(_x.name, _x.content)
                })
            }
            console.log(item)
            return true
        } catch (err) {
            return false

        }
    }

    public updateItemById = async (listName: string, obj: any, attachments: any, id: number): Promise<boolean> => {
        try {
            const item: any = await this._sp.web.lists.getByTitle(listName).items.getById(id).update(obj);
            if (attachments?.length > 0) {

                const itemById: IItem = await this._sp.web.lists.getByTitle(listName).items.getById(id);
                attachments?.map(async (_x: any) => {
                    await itemById.attachmentFiles.add(_x.name, _x.content)
                })
            }
            console.log(item)
            return true
        } catch (err) {
            return false

        }
    }


    public getItemById = async (listName: string, itemId: number, selectFields: string, expand: string) => {
        const items = await this._sp.web.lists.getByTitle(listName).items.getById(itemId).select(`*,${selectFields}`).expand(expand)();
        const info: any[] = await this._sp.web.lists.getByTitle(listName).items.getById(itemId).attachmentFiles();
        const files: any = []
        if (info && info.length > 0) {
            info?.map((_x: any, index: number) => {
                files.push({
                    "name": _x.FileName,
                    "content": null,
                    "index": index,
                    "fileUrl": "",
                    "ServerRelativeUrl": _x.ServerRelativeUrl,
                    "isExists": true,
                    "Modified": null,
                    "isSelected": false
                })
            })
        }
        return { items, files };
    }

    public getCustomerDrpDwnOption = async (filterColumn: any): Promise<any[]> => {
        const options: any = []
        // const items = await this._sp.web.lists.getByTitle("CustomerMaster").items.filter(`field_1 eq '${filterColumn}'`).select(`Title,ID`)();
        const items = await this._sp.web.lists.getByTitle("Customer Master").items.select(`*,Title,ID`).top(5000)();
        if (items) {
            items.map(_x => {
                options.push({
                    key: _x.ID,
                    text: _x.Title
                })
            })

        }
        return options;
    }

    /* Get userId  */
    public getUserId = async (username: string) => {
        const result = await this._sp.web.ensureUser(username);
        return result.Id
    }

    public getPreSaleSPoOptions = async (opptyType: string) => {
        const _DrpOptions: any = [];
        let filterQuery = ``;
        if (opptyType) {
            filterQuery = `Title eq '${opptyType}'`
        }

        const response = await this._sp.web.lists.getByTitle("Pre-sales Spoc & Opportunity Type Configuration").items.filter(filterQuery).select(`*,PreSalesSpoc/EMail,PreSalesSpoc/Title,PreSalesSpoc/Id`).expand("PreSalesSpoc")();
        if (response) {
            response.map(_y => {
                const PreSalesSpoc = _y.PreSalesSpoc || [];
                if (PreSalesSpoc) {
                    PreSalesSpoc.map((_x: { Id: any; Title: any; }) => {
                        _DrpOptions.push({
                            key: _x.Id, text: _x.Title
                        })
                    })
                }
            })

        }
        return _DrpOptions;
    }

}