import {
  IBizpNonRecurrence,
  IBizpRecurrence,
  IBizpEntryTypeEnum,
 } from '../../../../shared/IBizpSharedInterface';
export interface IBizpRecurrentEventProps {
  event?: IBizpNonRecurrence;
  eventSeries?: IBizpRecurrence;
  entryType:IBizpEntryTypeEnum;
  infoRequest:boolean;
  returnInfo: (returnData:IBizpRecurrence,initializing:boolean) => void;
  startDateChange: (newDate: Date) => void;
}

