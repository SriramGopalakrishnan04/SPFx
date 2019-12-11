import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';

import * as strings from 'SharePointSiteRequestFormWebPartStrings';
import SharePointSiteRequestForm from './components/SharePointSiteRequestForm';
import { ISharePointSiteRequestFormProps } from './components/ISharePointSiteRequestFormProps';

// Needed for IE Support
import "@pnp/polyfill-ie11";

export interface ISharePointSiteRequestFormWebPartProps {
  listName: string;
}

export default class SharePointSiteRequestFormWebPart extends BaseClientSideWebPart<ISharePointSiteRequestFormWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ISharePointSiteRequestFormProps> = React.createElement(
      SharePointSiteRequestForm,
      {
        listName: this.properties.listName,
        webpartContext: this.context
      }
    );
    
    ReactDom.render(element, this.domElement);
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
                PropertyPaneTextField('listName', {
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
