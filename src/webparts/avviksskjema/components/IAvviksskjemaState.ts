export interface IAvviksskjemaState {
  // common
  category: string; // choice
  incidentDate?: Date;
  incidentLocation: string; // 200 characters
  incidentDescription: string; // 32768 characters
  incidentConsecquences: string; // 32768 characters
  incidentCause: string; // 32768 characters
  suggestedActions: string; // 32768 characters
  priority: string;

  // personopplysninger på avveie
  peopleInvolved: string; // 32768 characters
  incidentToDate?: Date;
  incidentFoundDateTime?: Date;
  incidentMainCause: string; // choice
  relationsForPeopleInvolved: string; // choice
  relationsForPeopleInvolvedOther: string;

  // brudd på policy
  isRelatedToSecurityLaw: string;
  involvedUnit: string;

  // mangel på policy/security exception
  involved: string;
  missingPolicy: string;
  resultNeeds: string;
  suggestedResolution: string;

  // non-form fields
  responseID?: string;
  hasError: boolean;
  errorMessage?: string;
  errorCode?: string;
  sending: boolean;
}

export const DefatultState: IAvviksskjemaState = {
  category: '',
  incidentDate: undefined,
  incidentLocation: '',
  incidentDescription: '',
  incidentConsecquences: '',
  incidentCause: '',
  suggestedActions: '',
  priority: '',
  peopleInvolved: '',
  incidentToDate: undefined,
  incidentFoundDateTime: undefined,
  incidentMainCause: '',
  relationsForPeopleInvolved: '',
  relationsForPeopleInvolvedOther: '',
  isRelatedToSecurityLaw: '',
  involvedUnit: '',
  involved: '',
  missingPolicy: '',
  resultNeeds: '',
  suggestedResolution: '',
  responseID: undefined,
  hasError: false,
  errorCode: undefined,
  errorMessage: undefined,
  sending: false,
};
