import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';

import * as strings from 'BranchStylesApplicationCustomizerStrings';

const LOG_SOURCE: string = 'BranchStylesApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it
 */
export interface IBranchStylesApplicationCustomizerProperties {
  // This is an example; replace with your own property
  cssurl: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class BranchStylesApplicationCustomizer
  extends BaseApplicationCustomizer<IBranchStylesApplicationCustomizerProperties> {

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    //Inject the style element
      
    const head: any = document.getElementsByTagName("head")[0] || document.documentElement;

    //Beginning of adding custom css file
    // const cssURL: string = this.properties.cssurl;  
    // let customStyle: HTMLLinkElement = document.createElement("link");
    // customStyle.href = "/Style%20Library/BranchStyles.css";
    // customStyle.rel = "stylesheet";
    // customStyle.type = "text/css";
    // console.log("cssurl:"+this.properties.cssurl);
    // console.log("BranchStyles injected the style:"+customStyle.outerHTML);
    // head.insertAdjacentElement("beforeEnd", customStyle);
    //End of adding custom css file
    
    let styleElement = document.createElement("style");
    //customStyle.innerHTML="'button[name='Share'] {display: none;} button[data-automationid='ShareSiteButton'] {display: none;}'";
    var customStyle="button[name='Share'] {display: none;} button[data-automationid='ShareSiteButton'] {display: none;}";
    styleElement.type="text/css";
    styleElement.appendChild(document.createTextNode(customStyle))
    console.log(styleElement.innerHTML);
    head.insertAdjacentElement("beforeEnd", styleElement);

    return Promise.resolve();
  }
}
