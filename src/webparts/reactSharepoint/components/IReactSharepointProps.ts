import { SPHttpClient } from '@microsoft/sp-http'; 
import { WebPartContext } from '@microsoft/sp-webpart-base';
export interface IReactSharepointProps {
  description: string;
  context: WebPartContext;
  spHttpClient: SPHttpClient;  
  siteUrl: string; 
}
