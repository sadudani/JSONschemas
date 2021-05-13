import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IBizpCalendarEventsDisplayProps {
  siteUrl: string;
  list: string;
  context:WebPartContext;
  refresh?:boolean;
  daysInFuture: number;
  daysInPast: number;
}
