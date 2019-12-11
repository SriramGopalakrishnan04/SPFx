import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'ConnectedSource2WebPartStrings';
import ConnectedSource2 from './components/ConnectedSource2';
import { IConnectedSource2Props } from './components/IConnectedSource2Props';

export interface IConnectedSource2WebPartProps {
  description: string;
}

export default class ConnectedSource2WebPart extends BaseClientSideWebPart<IConnectedSource2WebPartProps> {

  public render(): void {
    const element: React.ReactElement<IConnectedSource2Props > = React.createElement(
      ConnectedSource2,
      {
        description: this.properties.description
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
