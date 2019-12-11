import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'SpFxReactFormPnPControlsWebPartStrings';
import SpFxReactFormPnPControls from './components/SpFxReactFormPnPControls';
import { ISpFxReactFormPnPControlsProps } from './components/ISpFxReactFormPnPControlsProps';




export interface ISpFxReactFormPnPControlsWebPartProps {
  description: string;
}

export default class SpFxReactFormPnPControlsWebPart extends BaseClientSideWebPart<ISpFxReactFormPnPControlsWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ISpFxReactFormPnPControlsProps > = React.createElement(
      SpFxReactFormPnPControls,
      {
        context: this.context,
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
