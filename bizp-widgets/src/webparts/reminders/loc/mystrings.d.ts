declare interface IRemindersWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  CalendarSelectionGroupName: string;
  HelpMenuLabel:string;
  FeedbackMenuLabel:string;
  RefreshMenuLabel:string;
}

declare module 'RemindersWebPartStrings' {
  const strings: IRemindersWebPartStrings;
  export = strings;
}
