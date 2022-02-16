import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IAvviksskjemaProps {
  context: WebPartContext;
  azureFunctionUrl: string;
  azureFunctionCode: string;
}
