import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'ConnectedConsumerWebPartStrings';
import ConnectedConsumer from './components/ConnectedConsumer';
import { IConnectedConsumerProps } from './components/IConnectedConsumerProps';

export interface IConnectedConsumerWebPartProps {
  description: string;
}

export default class ConnectedConsumerWebPart extends BaseClientSideWebPart<IConnectedConsumerWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IConnectedConsumerProps > = React.createElement(
      ConnectedConsumer,
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
