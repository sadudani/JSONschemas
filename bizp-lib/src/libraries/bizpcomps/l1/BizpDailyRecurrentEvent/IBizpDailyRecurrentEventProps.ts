import { IBizpDailyRecurrence,IBizpNonRecurrence, IBizpRecurrence,IBizpEntryTypeEnum} from '../../../../shared/IBizpSharedInterface';
export interface IBizpDailyRecurrentEventProps {
  event?: IBizpNonRecurrence;
  eventSeries?: IBizpRecurrence;
  entryType:IBizpEntryTypeEnum;
  infoRequest: boolean;
  returnInfo: (returnData:IBizpDailyRecurrence) => void;
  startDateChange: (newDate: Date) => void;
}
