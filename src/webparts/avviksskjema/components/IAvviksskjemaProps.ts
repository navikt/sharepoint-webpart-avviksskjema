import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IAvviksskjemaProps {
  context: WebPartContext;
  salesforceUrl: string;
  salesforceToken: string;
}
