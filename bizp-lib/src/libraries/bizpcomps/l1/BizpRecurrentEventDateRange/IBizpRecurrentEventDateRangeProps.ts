import { IBizpNonRecurrence,IBizpRecurrenceDateRange,IBizpEntryTypeEnum} from '../../../../shared/IBizpSharedInterface';
import { WebPartContext } from "@microsoft/sp-webpart-base";
export interface IBizpRecurrentEventDateRangeProps {
  event?: IBizpNonRecurrence;
  dateRange?: IBizpRecurrenceDateRange;
  entryType:IBizpEntryTypeEnum;
  infoRequest: boolean;
  returnInfo: (returnData:IBizpRecurrenceDateRange) => void;
  startDateChange: (newDate: Date) => void;
}
