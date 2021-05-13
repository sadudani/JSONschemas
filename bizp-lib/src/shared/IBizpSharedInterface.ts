import * as strings from 'BizpcompsLibraryStrings';
export interface IBizpListSpec {
  listName: string;
  siteURL: string;
  viewName: string;
}
export interface IBizpUserPermissions {
  hasPermissionAdd: boolean;
  hasPermissionEdit: boolean;
  hasPermissionDelete: boolean;
  hasPermissionView: boolean;
}

export type IBizpColumnTypes =  'BooleanCheck'|'Integer'|'Number'|'Boolean'|'String'|'Date'|'ShortDate'|'Time'|'Currency';
export type IBizpCalendarEventTypes =  'Reminder'|'Calendar';
export type ListStateType =  {
  listItems: any[];
  selectedItem: any;
};
export type IBizpPeriodOptionTypes =
'Last 30 days' | 'Last 90 days' | 'Last 6 months' | 'Last year' |
'Next 30 days' | 'Next 90 days';

export type IBizpOrderTypes = 'asc' | 'desc';

export interface IBizpOrderOptionTypes {
  field: string;
  order: IBizpOrderTypes;
}

export interface IBizpGridFieldSpec {
  name: string; // field name in the data bound to the column
  title: string; // corresponding title for the grid column
  type: IBizpColumnTypes; // type for styling and validation
}
export interface IBizpGridDisplaySpec {
  pageable: boolean;
  sortable: boolean;
}
export interface IBizpGridSpec {
  displaySettings:IBizpGridDisplaySpec;
  fields: IBizpGridFieldSpec[];
}
export interface IBizpWindowSpec {
  windowHeight:any;
  windowWidth: any;
  webpartTitle:string; // "My Reminders";
  windowId: string; // "manageEntries"; // Div id for manageAllReminder k-window modal
  windowTitle: string;
  usageView: string; // 'All Items View';
}
export interface IBizpContentEntryFormSpec {
  windowHeight:any;
  windowWidth: any;
  windowId: string; // "manageEntries"; // Div id for manageAllReminder k-window modal
  windowTitle: string; // "Reminder"
  usageView: string; // 'All Items View';
}
export interface IBizpMenuOptions {
  key:string;
  text:string;
  iconName: string;
}

export interface ICalendarEventCategory {
  category: string;
  color: string;
}
export interface IBizpCalendarRequest {
  includeRecurringEvents:boolean;
  listName: string;
  viewName?: string;
  siteURL: string;
  eventStartDate:string;
  eventEndDate:string;
}
export interface IUserPermissions {
  hasPermissionAdd: boolean;
  hasPermissionEdit: boolean;
  hasPermissionDelete: boolean;
  hasPermissionView: boolean;
}

export enum  IBizpCRUDEnum {
  add=1,
  edit=2,
  delete=3,
  view=4
}

export enum  IBizpEntryTypeEnum {
  newEvent=1,
  editEvent=2,
  editEventFromSeries=3,
  editSeries=4,
  viewEvent=5
}

export interface IBizpDaysCheck {
  sunday:boolean;
  monday:boolean;
  tuesday:boolean;
  wednesday:boolean;
  thursday:boolean;
  friday:boolean;
  saturday:boolean;
}

export interface IBizpEventDataSpec {
  id?: number;
  ID?: number;
  title: string;
  description: any;
  startDate: string;
  endDate?: string;
  fAllDayEvent?: boolean;
  duration?: number;
  recurrenceID?: string;
  recurrenceData?: string;
  fRecurrence?: string;
  category?: string;
  eventType?: string;
  attendes?: number[];
  location?: string;
  geolocation?: { Longitude: number, Latitude: number };
  color?: string;
  ownerInitial?: string;
  ownerPhoto?: string;
  ownerEmail?: string;
  ownerName?: string;
  UID?: string;
  masterSeriesItemID?: string;
}

export interface IBizpRecurrenceDateRange {
  startDate: string|Date;
  endDate: string|Date;
  frequency: string;
  option: string; //'noDate','endDate','endAfter'
}
export interface IBizpDailyRecurrence {
  dateRangeInfo: IBizpRecurrenceDateRange;
  frequency: string;
  pattern:string; // 'every', 'everyweekday'
}
export interface IBizpWeeklyRecurrence {
  dateRangeInfo: IBizpRecurrenceDateRange;
  frequency: string;
  daysOption: IBizpDaysCheck;
}
export interface IBizpMonthlyRecurrence {
  dateRangeInfo: IBizpRecurrenceDateRange;
  frequency: string;
  pattern:string; // dayOfMonth or OrderInMonth
  dayOfMonth: string; // 1 or 2 or ... or 31
  orderInMonth:string; // first, second,third,fourth or last
  dayOption:string; // day, weekday, weekendday, sunday, monday ... or saturday
}
export interface IBizpYearlyRecurrence {
  dateRangeInfo: IBizpRecurrenceDateRange;
  pattern:string; // byDay, byDayPattern
  month: string; // 1..12 (January, ....,December)
  dayOfMonth: string; // 1 or 2 or ... or 31
  orderInMonth:string; // first, second,third,fourth or last
  dayOption:string; // day, weekday, weekendday, sunday, monday ... or saturday
  dayPatternMonth:string; // 1..12 (January, ....,December)
}
export interface IBizpRecurrence {
  rule:string;
  dailyRecurrence?: IBizpDailyRecurrence;
  weeklyRecurrence?: IBizpWeeklyRecurrence;
  monthlyRecurrence?: IBizpMonthlyRecurrence;
  yearlyRecurrence?: IBizpYearlyRecurrence;
}
export interface IBizpNonRecurrence {
  title: string;
  description: any;
  startDate: Date;
  endDate?: Date;
  eventType: string;
  fAllDayEvent?: boolean;
  duration?: number;
  fRecurrence?: string;
  category?: string;
  attendees?: number[];
  location?: string;
  geolocation?: { Longitude: number, Latitude: number };
  color?: string;
}

