//import { IWebPartContext } from '@microsoft/sp-webpart-base';
import { WebPartContext } from "@microsoft/sp-webpart-base";
export interface IAdvanceSearchProps {
  description: string;
  urlSite: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
  
}
