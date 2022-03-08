export interface IAvviksskjemaState {
  // common
  category: string; // choice
  incidentDate?: Date;
  incidentLocation: string; // 200 characters
  incidentDescription: string; // 32768 characters
  suggestedActions: string; // 32768 characters

  personalInfoLost: boolean;
  categoryOther: string;

  // personopplysninger på avveie
  peopleInvolved: string; // 32768 characters
  incidentToDate?: Date;
  incidentMainCause: string; // choice
  relationsForPeopleInvolved: string; // choice
  relationsForPeopleInvolvedOther: string;

  // non-salesforce fields
  personvernCategory: string;
  responseID?: string;
  hasError: boolean;
  errorMessage?: string;
  errorCode?: string;
  sending: boolean;
  debug: boolean;
}

export const DefatultState: IAvviksskjemaState = {
  category: '',
  categoryOther: '',
  incidentDate: new Date(),
  incidentLocation: '',
  incidentDescription: '',
  suggestedActions: '',
  peopleInvolved: '',
  incidentToDate: new Date(),
  incidentMainCause: '',
  relationsForPeopleInvolved: '',
  relationsForPeopleInvolvedOther: '',
  personvernCategory: '',
  responseID: undefined,
  hasError: false,
  errorCode: undefined,
  errorMessage: undefined,
  sending: false,
  personalInfoLost: false,
  debug: false,
};
