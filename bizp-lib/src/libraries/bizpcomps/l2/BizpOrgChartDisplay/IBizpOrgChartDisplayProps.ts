import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IReadonlyTheme } from '@microsoft/sp-component-base';

export interface IBizpOrgChartDisplayProps {
  list?: string;
  context:WebPartContext;
  theme:IReadonlyTheme;
  refresh?:boolean;
  layout?: number;
}

export interface IBizpUserData {
  displayName: string;
  id: string;
  jobTitle: string;
  mail: string;
  mobilePhone: string;
  officeLocation: string;
  preferredLanguage: string;
  surname: string;
  userPrincipalName: string;
  manager: {id:string};
}
export interface IBizpOrgHierarchyData {
  children:IBizpOrgHierarchyData[];
  data:IBizpUserData;
}
