import * as React from 'react';
import styles from './MyReactForm.module.scss';
import { IMyReactFormProps } from './IMyReactFormProps';
import { IMyReactFormState } from './IMyReactFormState'
import { escape } from '@microsoft/sp-lodash-subset';
import {TextField} from 'office-ui-fabric-react/lib/TextField';
import {Button} from 'office-ui-fabric-react/lib/Button';
import { default as pnp, ItemAddResult, PnPClientStorage } from "sp-pnp-js";
import { FocusZone } from 'office-ui-fabric-react/lib/FocusZone';
import { List } from 'office-ui-fabric-react/lib/List';
import { DetailsList, DetailsListLayoutMode, Selection, IColumn } from 'office-ui-fabric-react/lib/DetailsList';
import { MarqueeSelection } from 'office-ui-fabric-react/lib/MarqueeSelection';
import { Fabric } from 'office-ui-fabric-react/lib/Fabric';
import { IRectangle } from 'office-ui-fabric-react/lib/Utilities';
import { ITheme, getTheme, mergeStyleSets } from 'office-ui-fabric-react/lib/Styling';


export default class MyReactForm extends React.Component<IMyReactFormProps, IMyReactFormState> {
  constructor(props) {
    super(props);
    this.state={ name:"", primaryOwner:"", items:[] };
    //this.createItem=this.createItem.bind(this);
    this.handleNameChange = this.handleNameChange.bind(this);
    this.handlePnPCreateItem = this.handlePnPCreateItem.bind(this);
    this.handlePnPReadItem=this.handlePnPReadItem.bind(this);
  }
  
  public render(): React.ReactElement<IMyReactFormProps> {
    const name=this.state;
    return (
      <form>
      <div className={ styles.myReactForm }>
        
              <div className={styles.container}>
                <div className={styles.row}>
                <div className={styles.column}>
              <TextField label="Name:" value={this.state.name} onChanged={this.handleNameChange} />
              </div>
              <div className={styles.column}>
              <Button onClick={ this.handlePnPCreateItem } text="Create Item" />
              </div>
              <div className={styles.column}>
              <Button onClick={ this.handlePnPReadItem } text="Read Item" />
              
        <FocusZone>
        <List
          items={this.state.items}
          onRenderCell={this._onRenderCell}
        />
        </FocusZone>

        
        </div>
      
              </div>
              </div>              
            </div>         
      </form>
    );
  }

  private handlePnPCreateItem(): void {
    if(this.state.name==""){
   alert("Name cannot be blank");
    }
    else {
    pnp.sp.web.lists.getByTitle("Test").items.add({
      'Title': this.state.name
    });
  }

  }

  private handleNameChange(value: string): void {
    return this.setState({
      name: value
    });
  }

  private handlePnPReadItem(): void {
    // pnp.sp.web.lists.getByTitle("Test").items.get().then(
    //   (items: any[]) => {items.forEach((item)=>{this})}
    //   )
    pnp.sp.web.lists.getByTitle("Test").items.get().then(
         (splistitems: any[]) => {this.setState({items:splistitems});}
       )

    
  }

  private _onRenderCell = (item: any): JSX.Element => {
    return (
      <div>
            <span>{item['Title']}</span>
            <span>{item['Age']}</span>
            </div>
         
     )
  }


}
