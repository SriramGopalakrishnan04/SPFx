import * as React from 'react';
import styles from './ConnectedSource2.module.scss';
import { IConnectedSource2Props } from './IConnectedSource2Props';
import { escape } from '@microsoft/sp-lodash-subset';
import { ListView, IViewField, SelectionMode, GroupOrder, IGrouping } from "@pnp/spfx-controls-react/lib/ListView";
import { sp, Items } from '@pnp/sp';
import { string } from 'prop-types';
import { IConnectedSource2WebPartProps } from '../ConnectedSource2WebPart';

export interface ITeamSiteRequest {
  Title: string;
  Description: string;

}
export interface IConnectedSource2State {
  SiteRequests: ITeamSiteRequest[];
}

export default class ConnectedSource2 extends React.Component<IConnectedSource2Props, IConnectedSource2State> {

  constructor(props: IConnectedSource2WebPartProps){
    super(props);
    this.state={SiteRequests: []};
  }
  

  public componentDidMount() {
    
    sp.web.lists.getByTitle("Project Requests").items.select("Title","Description").get()
      .then((response)=> {
        console.log(response); 
        this.setState({SiteRequests:response});
      });
  }

  private _getSelection(items: any[]) {
    console.log('Selected items:', items);
    
  }

  public render(): React.ReactElement<IConnectedSource2Props> {
    const viewFields: IViewField[] = [
      {
        name: 'Title',
        displayName: 'Project Request',
        sorting: true,
        maxWidth: 400
      },
      {
        name: 'Description',
        displayName: 'Description',
        sorting: true,
        maxWidth: 80
      }
    ];
    return (
      
      <div className={styles.connectedSource2}>
      
        <ListView
          items={this.state.SiteRequests}
          viewFields={viewFields}
          iconFieldName="ServerRelativeUrl"
          compact={true}
          selectionMode={SelectionMode.single}
          selection={this._getSelection}
          showFilter={true}
          filterPlaceHolder="Search..."
        />
      </div>
    );
  }
}
