import * as React from 'react';
import styles from './ConnectedConsumer.module.scss';
import { IConnectedConsumerProps } from './IConnectedConsumerProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { IConnectedConsumerWebPartProps } from '../ConnectedConsumerWebPart';
import { sp, Items } from '@pnp/sp';

export interface IConnectedConsumerState {
  ProjectDetails: IProjectDetails[];
}

export interface IProjectDetails {
  Title: string;
  ApprovedDate: string;
  Status: string;
}

export default class ConnectedConsumer extends React.Component<IConnectedConsumerProps, IConnectedConsumerState> {

  constructor(props: IConnectedConsumerWebPartProps) {
    super(props);
    this.state = { ProjectDetails: [{ Title: "", ApprovedDate: "", Status: "" }] };
  }

  public componentDidMount() {
     sp.web.lists.getByTitle("Project Details").items.select("Title","ApprovedDate","Status").getAll()
     .then((response:IProjectDetails[])=> {
       console.log(response[0]); 
       this.setState({ProjectDetails:response});
     });
  }

  public render(): React.ReactElement<IConnectedConsumerProps> {
    return (
      <div className={styles.connectedConsumer}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div>Project Title: {this.state.ProjectDetails[0].Title} </div>
            <div>Approved Date: {this.state.ProjectDetails[0].ApprovedDate}</div>
            <div>Status: {this.state.ProjectDetails[0].Status}</div>
          </div>
        </div>
      </div>
    );
  }
}
