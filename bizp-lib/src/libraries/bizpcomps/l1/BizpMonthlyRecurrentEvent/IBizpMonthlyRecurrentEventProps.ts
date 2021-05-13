import { IBizpNonRecurrence, IBizpRecurrence, IBizpMonthlyRecurrence,IBizpEntryTypeEnum} from '../../../../shared/IBizpSharedInterface';
export interface IBizpMonthlyRecurrentEventProps {
  event?: IBizpNonRecurrence;
  eventSeries?: IBizpRecurrence;
  entryType:IBizpEntryTypeEnum;
  infoRequest: boolean;
  returnInfo: (returnData:IBizpMonthlyRecurrence) => void;
  startDateChange: (newDate: Date) => void;
}

