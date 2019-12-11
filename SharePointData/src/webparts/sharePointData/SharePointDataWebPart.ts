import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'SharePointDataWebPartStrings';
import SharePointData from './components/SharePointData';
import { ISharePointDataProps } from './components/ISharePointDataProps';
import {SPHttpClient, SPHttpClientConfiguration} from '@microsoft/sp-http';
import { IProjectRequestItem } from './models/ProjectRequests';
import {IProjectRequestItems} from './models/ProjectRequests';
import styles from './components/SharePointData.module.scss';

export interface ISharePointDataWebPartProps {
  description: string;
}

export default class SharePointDataWebPart extends BaseClientSideWebPart<ISharePointDataWebPartProps> {

  private projects: IProjectRequestItem[] = [];
  public render(): void {
    const element: React.ReactElement<ISharePointDataProps > = React.createElement(
      SharePointData,
      {
        description: this.properties.description        
      }
    );
    
    //this.renderListAsync();
    this.renderListItemAsync();
    ReactDom.render(element, this.domElement);
  }

  private getListData(): Promise<IProjectRequestItems>
  {
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl+"/_api/web/lists/getByTitle('Project Requests')/Items?$select=Id,Title",SPHttpClient.configurations.v1)
    .then(      
      response=>{
      return response.json();
    });
  }

  private getListItemById():Promise<IProjectRequestItem>
  {
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl+"/_api/web/lists/getByTitle('Project Requests')/items(2)?$select=Id,Title",SPHttpClient.configurations.v1)
    .then(response => {
      return response.json();
    });
    
  }

  private renderListData(items:IProjectRequestItem[]):void {
    let listdata: string = '';
    items.forEach((item:IProjectRequestItem) => {
      listdata+=`<ul">
      <li>
        <span class="ms-font-l">${item.Id}</span>
        <span class="ms-font-l">${item.Title}</span>
      </li>
      </ul>`;
      const listContainer: Element = this.domElement.querySelector('#spListContainer');
      
      listContainer.innerHTML = listdata;
    });
  }

  private renderListItem(item:IProjectRequestItem):void {
    let listdata: string = '';
    
      listdata+=`<ul">
      <li className= ${styles.row}>
        <span class="ms-font-l">${item.Id}</span>
        <span class="ms-font-l">${item.Title}</span>
      </li>
      </ul>`;
      const listContainer: Element = this.domElement.querySelector('#spListContainer');
      
      listContainer.innerHTML = listdata;
    
  }

  private renderListAsync(): void {
    this.getListData()
      .then((response) => {
        this.renderListData(response.value);
      });
  }

  private renderListItemAsync(): void {
    this.getListItemById()
    .then(item => {
      this.renderListItem(item);
    });       
      
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
