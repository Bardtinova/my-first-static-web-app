import { WebPartContext } from "@microsoft/sp-webpart-base";
//import { BaseComponentContext } from '@microsoft/sp-component-base';

 interface INewPartProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  selectedList: string;
  
}

export interface ITestProps extends INewPartProps {
  context: WebPartContext;
}
