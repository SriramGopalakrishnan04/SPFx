import * as React from 'react';
import styles from './ConnectedSource.module.scss';
import { IConnectedSourceProps } from './IConnectedSourceProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { ListView, IViewField, SelectionMode, GroupOrder, IGrouping } from "@pnp/spfx-controls-react/lib/ListView";
import { sp, Items } from '@pnp/sp';
import { string } from 'prop-types';
import { IConnectedSourceWebPartProps } from '../ConnectedSourceWebPart';

export interface ITeamSiteRequest {
  Title: string;
  Description: string;

}
export interface IConnectedSourceState {
  SiteRequests: ITeamSiteRequest[];
}

export default class ConnectedSource extends React.Component<IConnectedSourceProps, IConnectedSourceState> {

  constructor(props: IConnectedSourceWebPartProps){
    super(props);
    this.state={SiteRequests: []};
  }
  

  public componentDidMount() {
    
    sp.web.lists.getByTitle("Project Requests").items.select("Title","Description").get()
      .then((response)=> {
        console.log(response); 
        this.setState({SiteRequests:response});
      })
  }

  private _getSelection(items: any[]) {
    console.log('Selected items:', items[0].Title);
    
  }

  public render(): React.ReactElement<IConnectedSourceProps> {
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
      
      <div className={styles.connectedSource}>
      
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
