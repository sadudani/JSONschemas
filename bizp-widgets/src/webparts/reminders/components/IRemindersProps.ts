import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IRemindersProps {
  description: string;
  title: string;
  siteUrl: string;
  list: string;
  context:WebPartContext;
  daysInFuture: number;
  daysInPast: number;
}
