import { IBizpNonRecurrence, IBizpRecurrence, IBizpEventDataSpec,IBizpCRUDEnum } from '../../../../shared/IBizpSharedInterface';
export interface IBizpEventEntryProps {
  event?: IBizpEventDataSpec;
  panelMode: IBizpCRUDEnum;
  series?: boolean;
  eventSeries: IBizpRecurrence;
  onDismissEntry?: (refresh:boolean) => void;
  onSaveNewEntry?: (eventData: IBizpNonRecurrence,recurrentData:IBizpRecurrence) => void;
  showPanel: boolean;
  startDate?: Date;
  endDate?: Date;
  siteUrl: string;
  listId:string;
}

