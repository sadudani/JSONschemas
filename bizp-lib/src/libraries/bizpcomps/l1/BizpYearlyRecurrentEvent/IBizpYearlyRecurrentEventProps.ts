import { IBizpNonRecurrence,IBizpRecurrence,IBizpYearlyRecurrence,IBizpEntryTypeEnum } from '../../../../shared/IBizpSharedInterface';
export interface IBizpYearlyRecurrentEventProps {
  event?: IBizpNonRecurrence;
  eventSeries?: IBizpRecurrence;
  entryType:IBizpEntryTypeEnum;
  infoRequest: boolean;
  returnInfo: (returnData:IBizpYearlyRecurrence) => void;
  startDateChange: (newDate: Date) => void;
}
