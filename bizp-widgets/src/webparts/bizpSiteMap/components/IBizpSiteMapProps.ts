import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IReadonlyTheme } from "@microsoft/sp-component-base";

export interface IBizpSiteMapProps {
  description: string;
  title: string;
  siteUrl: string;
  list: string;
  context:WebPartContext;
  themeVariant: IReadonlyTheme;
}
