import * as React from 'react';
import styles from './SharePointData.module.scss';
import { ISharePointDataProps } from './ISharePointDataProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class SharePointData extends React.Component<ISharePointDataProps, {}> {
  public render(): React.ReactElement<ISharePointDataProps> {
    return (
      <div className={ styles.sharePointData }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>Welcome to SharePoint! v1.1</span>
              <p className={ styles.subTitle }>Customize SharePoint experiences using Web Parts.</p>
              <p className={ styles.description }>{escape(this.props.description)}</p>
              <div id="spListContainer">my data test</div>
              <a href="https://aka.ms/spfx" className={ styles.button }>
                <span className={ styles.label }>Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
