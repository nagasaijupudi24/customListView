import * as React from 'react';
import 'office-ui-fabric-core/dist/css/fabric.min.css';


import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
interface IXenpreSalesPeopelPickerProps {
    context: any;
    onGetPeoplePickerItems: any;
    filterUsers: any;
    opptyType:any
}







export default class XenCustomPeoplePicker extends React.Component<IXenpreSalesPeopelPickerProps> {
    protected presaleSPOC: any;
    constructor(props: IXenpreSalesPeopelPickerProps) {
        super(props);
    } 
    private clear=()=>{
    this.presaleSPOC.state.selectedPersons = [];
    this.presaleSPOC.onChange([]);
    }

    




    private _getPeoplePickerItemsSingleUser = (items: any[]) => {
        this.props.onGetPeoplePickerItems("PresalesSPOCId", items || [])
    }

    private _filterUserInPeoplePicker = (result: any) => {
        const { filterUsers } = this.props;
        if (filterUsers.length > 0) {
            return result.filter((p: any) => filterUsers.includes(p["secondaryText"].toLowerCase()));
        } else {
            return result
        }
    }

    public getRender=()=>{
        return(    <PeoplePicker
                context={this.props.context}
                personSelectionLimit={1}
                groupName={""} // Leave this blank in case you want to filter from all users
                showtooltip={true}
                disabled={false}
                onChange={this._getPeoplePickerItemsSingleUser.bind(this)}
                showHiddenInUI={false}
                ensureUser={true}
                resultFilter={(result) => this._filterUserInPeoplePicker(result)}
                principalTypes={[PrincipalType.User]}
                resolveDelay={1000} 
                  ref={c => (this.presaleSPOC = c)}/>)
    }

    public componentDidUpdate(prevProps: Readonly<IXenpreSalesPeopelPickerProps>, prevState: Readonly<{}>, snapshot?: any): void {
        if(prevProps.opptyType !== this.props.opptyType){
            console.log("test")
            this.getRender();
            this.clear();
        }
        
    }


    public render(): React.ReactElement<IXenpreSalesPeopelPickerProps> {
        return (
<div>{this.getRender()}</div>
        
        );
    }
}