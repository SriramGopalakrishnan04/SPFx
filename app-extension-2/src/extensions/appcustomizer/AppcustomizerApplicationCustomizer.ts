import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';

import * as strings from 'AppcustomizerApplicationCustomizerStrings';

const LOG_SOURCE: string = 'AppcustomizerApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IAppcustomizerApplicationCustomizerProperties {
  // This is an example; replace with your own property
  cssUrl: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class AppcustomizerApplicationCustomizer
  extends BaseApplicationCustomizer<IAppcustomizerApplicationCustomizerProperties> {

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    const cssurl: string = this.properties.cssUrl;
    Dialog.alert(`1stalert ${strings.Title}:\n\n${cssurl}`);

     if (cssurl) {
       // inject the style sheet
       
       const head: any = document.getElementsByTagName("head")[0] || document.documentElement;
       let customStyle: HTMLLinkElement = document.createElement("link");
       customStyle.href = cssurl;
       customStyle.rel = "stylesheet";
       customStyle.type = "text/css";
       Dialog.alert(`insideifalert ${strings.Title}:\n\n${customStyle}`);
       head.insertAdjacentElement("beforeEnd", customStyle);
     }

     //Dialog.alert(`2ndalert ${strings.Title}:\n\n${cssurl}`);

    return Promise.resolve();
  }
}
