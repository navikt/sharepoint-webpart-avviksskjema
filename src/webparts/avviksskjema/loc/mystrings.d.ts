
declare interface IAvviksskjemaWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  IncidentCategoryPrivacy: string;
  IncidentCategorySecurity: string;
  IncidentCategoryHSE: string;
  IncidentCategoryOther: string;
  DateStrings: {
    goToToday: string,
    days: string[],
    shortDays: string[],
    months: string[],
    shortMonths: string[],
    dateLabel: string,
    timeLabel: string,
    timeSeparator: string,
    textErrorMessage: string,
  };
}

declare module 'AvviksskjemaWebPartStrings' {
  const strings: IAvviksskjemaWebPartStrings;
  export = strings;
}
