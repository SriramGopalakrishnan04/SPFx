import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/components/Dropdown';
 
export interface IReactSpFxPnP {
    selectedItems: any[];
    name: string; 
    description: string; 
    dpselectedItem?: { key: string | number | undefined };
    termKey?: string | number;
    dpselectedItems: IDropdownOption[];
    disableToggle: boolean;
    defaultChecked: boolean;
    pplPickerType:string;
    userManagerIDs: number[];
    hideDialog: boolean;
    status: string;
    isChecked: boolean;
    showPanel: boolean;
    required:string;
    onSubmission:boolean;
    termnCond:boolean;
}