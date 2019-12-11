import * as React from 'react';
import styles from './ConnectedWebPartConsumer.module.scss';
import { IConnectedWebPartConsumerProps } from './IConnectedWebPartConsumerProps';
import { escape } from '@microsoft/sp-lodash-subset';
import {sp, Item} from '@pnp/sp';

export interface IProjectDetails {
  Title: string;
  Status: string;
  ApprovedDate: string;
}

export interface IConnectedConsumerState {
  ProjectDetails: IProjectDetails;
}

export default class ConnectedWebPartConsumer extends React.Component<IConnectedWebPartConsumerProps, IConnectedConsumerState> {

  constructor(props: IConnectedWebPartConsumerProps){
    super(props);
    this.state={ProjectDetails: {Title:'', Status:'',ApprovedDate:''}};
  }

  public componentDidMount() {
    
    sp.web.lists.getByTitle("Project Details").items.select("Title","Status","ApprovedDate").getById(1).get().then(
      (response) => {
        this.setState({ProjectDetails:response})
      }
    )
    //console.log(projDetails);
    //this.setState({ProjectDetails:projDetails});
    // sp.web.lists.getByTitle("Project Details").items.select("Title","Status","ApprovedDate").getAll()
    //   .then((response)=> {
    //     console.log(response); 
    //     this.setState({ProjectDetails:response});
    //   });
  }


  public render(): React.ReactElement<IConnectedWebPartConsumerProps> {
    return (
      <div className={ styles.connectedWebPartConsumer }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>Project Details</span>
              {/* <div>Title: {this.state.ProjectDetails[0].Title}</div>
              <div>Status:{this.state.ProjectDetails[0].Status} </div>
              <div>Approved Date: {this.state.ProjectDetails[0].ApprovedDate} </div>               */}


              {/* {this.state.ProjectDetails.map((item)=> {
                console.log(item.Title);
                return ( 
                <div>{item.Title}
              <div>{item.Status}</div>
                <div>{item.ApprovedDate}</div> 
                </div>               
                );
              
               
              })} */}
  <div>{this.state.ProjectDetails.Title}</div> 
            </div>
          </div>
        </div>
      </div>
    );
  }
}
