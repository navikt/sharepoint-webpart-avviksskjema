
declare interface IAvviksskjemaWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  IncidentCategoryPrivacy: string;
  IncidentCategorySecurity: string;
  IncidentCategoryHSE: string;
  IncidentCategoryOther: string;

  PolicyBrudd: string;
  PolicyMangel: string;
  SecurityException: string;
  AndreHendelser: string;
  VoldTrusler: string;
  HMSavvik: string;
  HMSforbedringsforslag: string;
  PersonopplysningerPåAvveie: string;
  AnnetPersonvernrelatert: string;
  ManglendeBehandlingsgrunnlag: string;
  ManglendeIvaretagelseAvInnsynsretten: string;
  ManglerDatabehandleravtale: string;
  UgyldigSamtykke: string;
  IkkeFastsattLagringstider: string;
  BehandlingenIkkeRegistrertIBehandlingskatalogen: string;
  IkkeGjennomførtPVK: string;
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
