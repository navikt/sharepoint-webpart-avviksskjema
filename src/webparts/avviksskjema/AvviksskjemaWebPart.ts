import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneButton, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import Avviksskjema from './components/Avviksskjema';
import * as strings from 'AvviksskjemaWebPartStrings';

import { IAvviksskjemaProps } from './components/IAvviksskjemaProps';
import { sp } from '@pnp/sp';
import { HttpClient, IHttpClientOptions, HttpClientResponse } from '@microsoft/sp-http'; 

export interface IAvviksskjemaWebPartProps {
  azureFunctionUrl: string;
  azureFunctionCode: string;
}

export default class AvviksskjemaWebPart extends BaseClientSideWebPart<IAvviksskjemaWebPartProps> {

  public async onInit(): Promise<void> {
    await super.onInit();
    sp.setup(this.context);
  }

  public render(): void {
    const element: React.ReactElement<IAvviksskjemaProps> = React.createElement(
      Avviksskjema,
      {
        context: this.context,
        azureFunctionUrl: this.properties.azureFunctionUrl,
        azureFunctionCode: this.properties.azureFunctionCode,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: 'Nettdel for innsending av avviksskjemaer til Salesforce.'
          },
          groups: [
            {
              groupName: 'Tilkobling til Salesforce',
              groupFields: [
                PropertyPaneTextField('azureFunctionUrl', {
                  label: 'URL til Azure function',
                }),
                PropertyPaneTextField('azureFunctionCode', {
                  label: 'API-nøkkel',
                }),
              ]
            },
          ]
        }
      ]
    };
  }

}
