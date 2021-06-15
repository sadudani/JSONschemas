import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IReadonlyTheme } from '@microsoft/sp-component-base';

export interface IBizpSiteMapDisplayProps {
  siteUrl: string;
  list?: string;
  context:WebPartContext;
  refresh?:boolean;
  displayLibs: boolean;
  displayLists: boolean;
  layout: number;
  theme:IReadonlyTheme;
}

export interface IBizpSiteData {
  Rank?: string;
  DocId?: string;
  Title: string;
  Path: string;
  Description: string;
  ParentLink: string;
  SiteLogo?: string;
  WebTemplate: string;
  SiteId?: string;
  UniqueId?: string;
  WebId?: string;
  contentclass?: string;
  IsExternalContent?: string;
  ListId?: string;
  OriginalPath?: string;
  ParentSiteTitle?: string;
  ResultTypeIdList?: string;
  ResultTypeId?: string;
  RenderTemplateId?: string;
  piSearchResultId?: string;
  GeoLocationSource?: string;
  PartitionId?: string;
  UrlZone?: string;
  Culture?: string;
  anyLibs?:boolean;
  libsLoaded?:boolean;
}
export interface IBizpSiteHierarchyData {
  children:IBizpSiteHierarchyData[];
  data:IBizpSiteData;
}
