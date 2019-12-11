import {Dropdown,IDropdownOption} from 'office-ui-fabric-react/lib/components/Dropdown'

export interface ISPFxReactFormPnPControls {
    selectedItems: any[];
    name: string;
    description: string;
    dpselectedItem?: { key: string | number | undefined };
    termKey?: string | number;
    dpselectedItems: IDropdownOption[];
    disableToggle: boolean;
    defaultChecked: boolean;
    pplPickerType:string;
    userIDs: number[];
    userManagerIDs: number[];
    hideDialog: boolean;
    status: string;
    isChecked: boolean;
    showPanel: boolean;
    required:string;
    onSubmission:boolean;
    termnCond:boolean;
}