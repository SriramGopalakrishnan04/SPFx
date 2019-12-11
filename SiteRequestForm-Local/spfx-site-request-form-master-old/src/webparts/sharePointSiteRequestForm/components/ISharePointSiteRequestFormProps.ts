import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface ISharePointSiteRequestFormProps {
  listName: string;
  webpartContext: WebPartContext;
}