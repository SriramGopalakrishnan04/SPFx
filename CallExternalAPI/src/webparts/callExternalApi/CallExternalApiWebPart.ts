import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'CallExternalApiWebPartStrings';
import CallExternalApi from './components/CallExternalApi';
import { ICallExternalApiProps } from './components/ICallExternalApiProps';
import { HttpClient, HttpClientResponse, IHttpClientOptions } from '@microsoft/sp-http';

export interface ICallExternalApiWebPartProps {
  description: string;
}

export default class CallExternalApiWebPart extends BaseClientSideWebPart<ICallExternalApiWebPartProps> {

  public render(): void {
    // const apiURL="https://apistcld.sce.com/sce/stb/v1/googleassistant:getAccountBalance";
    // const requestHeaders: Headers = new Headers();
    // requestHeaders.append('X-IBM-Client-Id', 'b88f36c2-3073-4666-849f-e63cce009139');
    // requestHeaders.append('X-IBM-Client-Secret','D7hL1lT7cA0lG4eE4nF6wL2qD0eV7wB3qH2jA7mO1pM2uF3jH4');
    // const httpClientOptions: IHttpClientOptions = {
      
    //   headers: requestHeaders
    // };

    // this.context.httpClient
    // .get(apiURL, HttpClient.configurations.v1,httpClientOptions)
    // .then((res: HttpClientResponse): Promise<any> => {
    //   //alert(res.json());
    //   return res.json();      
    // })
    // .then((data: any): void => {
    //   // process your data here
    //   alert(JSON.stringify(data));
    // }, (err: any): void => {
    //   // handle error here
    //   alert(err);
    // });
    this.invokeIBMCloudAPI("https://52.157.22.95/sce/ut/v1/storage/tevfileshare?storageVersion=2018-03-28&storageService=f&storagePermission=rwdlc&storageExpiryDate=2021-09-10T02:33:49Z&storageStartDate=2019-09-27T18:33:49Z&storageProtocol=https&storageSignature=opNRbyIZLNW4ZvySJMGPQV0nZgKkvFKE1DGZSYICSKY%3D&storageResourceType=sco&directoryPath=MDHDTestDir100");
   //this.invokeIBMCloudAPI("https://apistcld.sce.com/sce/stb/v1/googleassistant:getAccountBalance");
    const element: React.ReactElement<ICallExternalApiProps > = React.createElement(
      CallExternalApi,
      {
        description: this.properties.description
      }
    );

    ReactDom.render(element, this.domElement);
  }

  private invokeIBMCloudAPI(apiUrl: string): void {
    try {
    alert('Beginning to call API');
    this.context.httpClient
  .get(apiUrl, HttpClient.configurations.v1,
    {
      headers: [
        ['accept', 'application/json'],
        ['Authorization', 'Basic YWExOTBmYjMtZDliMS00NjFlLTliNWItNjcwZWUxM2JhMjk5OlAyZEU4ZkcxdlUwakw2a1YwbFU4akc4bUs4alg1Y0M2bVgzdkcyb0Y0d1Izbko1bE0w'],
     ]
    })
  .then((res: HttpClientResponse): Promise<any> => {
    //alert(res.json());
    return res.json();
  })
  .then((response: any): void => {
  alert(JSON.stringify(response));
  },(err: any): void => {
       // handle error here
       alert(err);
     });
  
  
  

  }

catch(error) {
alert(error);
}
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
