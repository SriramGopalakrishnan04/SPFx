import * as React from 'react';
import { Dropdown, DropdownMenuItemType, IDropdownStyles, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { IStackTokens, Stack } from 'office-ui-fabric-react/lib/Stack';
//import {populateLifeExpectancy} from '../../services/sp-rest';
import { sp } from '@pnp/sp';
import { WebPartContext } from '@microsoft/sp-webpart-base';

interface IDropDownTemplateProps {
  label: string;
  placeHolder: string;
  webPartContext: WebPartContext;
  dataSource: string;
  dataText: string;
  dataValue?: string;
  required?: boolean;
}

var ddOptions: IDropdownOption[] = [{ key: "", text: "--Choose a Value--" }];

export interface IDropDownTemplateState {
  ddlOptions: IDropdownOption[];
  selectedItem?: { key: string | number | undefined };
}
const stackTokens: IStackTokens = { childrenGap: 20 };

const DropDownCss: Partial<IDropdownStyles> = {
  dropdown: { width: 500 }
};


class DropDownTemplate extends React.Component<IDropDownTemplateProps, {ddlOptions:IDropdownOption[] ; selectedKey: string}> {
  

  public constructor(props: IDropDownTemplateProps, state: IDropDownTemplateState) {
    super(props);
    this.state = {
      ddlOptions: [],
      selectedKey:""
    };
  }

  public componentDidMount(): IDropdownOption[] {
    sp.setup({
      spfxContext: this.props.webPartContext
    });
    if (this.props.dataText != '') {
      if (this.props.dataValue != '') {
        console.log("Data Value is not blank - " + this.props.dataValue);
        sp.web.lists.getByTitle(this.props.dataSource).items.select(this.props.dataText, this.props.dataValue).getAll().then((dataSet) => {
          (dataSet).map((data) =>
            ddOptions.push({ key: data[this.props.dataValue], text: data[this.props.dataText] })
          );
          this.setState({ ddlOptions: ddOptions });
          console.log(ddOptions);
        });

      } else {
        console.log("Data Value is blank");
        sp.web.lists.getByTitle(this.props.dataSource).items.select(this.props.dataText).getAll().then((dataSet) => {
          (dataSet).map((data) =>
            ddOptions.push({ key: data[this.props.dataText], text: data[this.props.dataText] })
          );
          this.setState({ ddlOptions: ddOptions });
          console.log(ddOptions);
        });

      }
    }
    return ddOptions;
  }


  private onChange = (
    ev: any,
    selectedOption: IDropdownOption | undefined
  ): void => {
    const selectedKey: string = selectedOption
      ? (selectedOption.key as string)
      : "";
    this.setState({ddlOptions:ddOptions, selectedKey:selectedKey});
  }


  public render() {

    return (
      <div>
      <Stack tokens={stackTokens}>
        <Dropdown
          placeholder={this.props.placeHolder}
          label={this.props.label}
          styles={DropDownCss}
          options={this.state.ddlOptions}
          required={this.props.required}
          selectedKey={this.state.selectedKey}
          onChange={this.onChange}
           />
      </Stack>
      Selected value:{this.state.selectedKey}
      </div>
    );
  }
}

export default DropDownTemplate;