import { ServiceScope } from '@microsoft/sp-core-library';  
import { WebPartContext } from '@microsoft/sp-webpart-base';
  
export interface IUserProfileViewerProps {  
  description: string;  
  userName: string;  
  serviceScope: ServiceScope;  
  context: WebPartContext;
}