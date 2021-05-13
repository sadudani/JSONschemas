import { IBizpNonRecurrence, IBizpRecurrence, IBizpWeeklyRecurrence,IBizpEntryTypeEnum} from '../../../../shared/IBizpSharedInterface';
export interface IBizpWeeklyRecurrentEventProps {
  event?: IBizpNonRecurrence;
  eventSeries?: IBizpRecurrence;
  entryType:IBizpEntryTypeEnum;
  infoRequest: boolean;
  returnInfo: (returnData:IBizpWeeklyRecurrence) => void;
  startDateChange: (newDate: Date) => void;
}

