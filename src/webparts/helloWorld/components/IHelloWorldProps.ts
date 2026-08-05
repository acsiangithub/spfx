//import { IWebPartContext } from '@microsoft/sp-webpart-base';
import { WebPartContext } from "@microsoft/sp-webpart-base";
export interface IHelloWorldProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
  
}
