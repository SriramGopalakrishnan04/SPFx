import * as React from 'react';
import { Dropdown, DropdownMenuItemType, IDropdownStyles, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { IStackTokens, Stack } from 'office-ui-fabric-react/lib/Stack';
//import {populateLifeExpectancy} from '../../services/sp-rest';
import {sp} from '@pnp/sp';
import { WebPartContext } from '@microsoft/sp-webpart-base';

interface IDropDownTemplateProps {
    label: string;
    placeHolder: string;
    webPartContext: WebPartContext;   
}

var ddOptions: IDropdownOption[]=[];
 
export interface IDropDownTemplateState{ 
  ddOptions: IDropdownOption[];
} 
const stackTokens: IStackTokens = { childrenGap: 20 };

const DropDownCss: Partial<IDropdownStyles> = {
    dropdown: { width: 500 }
  };


class DropDownTemplate extends React.Component<IDropDownTemplateProps, IDropDownTemplateState> {
    
public constructor(props: IDropDownTemplateProps, state: IDropDownTemplateState){ 
    super(props); 
    this.state = { 
      ddOptions: []
    }; 
  } 

   public  componentDidMount():IDropdownOption[] {
      sp.setup({
        spfxContext: this.props.webPartContext
      });  
      
      sp.web.lists.getByTitle("Test").items.select('Title').getAll().then((dataSet) => {
        (dataSet).map((data) =>
          ddOptions.push({key:data.Title, text:data.Title})
        )      
      this.setState({ddOptions});
      console.log(ddOptions);
      
      }); 
    return ddOptions;       
      }

   
    public render() {
        
        return(
            <Stack tokens={stackTokens}>
            <Dropdown placeholder={this.props.placeHolder} label={this.props.label} styles={DropDownCss} options={this.state.ddOptions} />
            </Stack>
        );
    }
}

export default DropDownTemplate;