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
  salesforceUrl: string;
  salesforceToken: string;
  accessTokenUrl: string;
  accessTokenUserName: string;
  accessTokenPassword: string;
  accessTokenClientID: string;
  accessTokenSecret: string;
  accessToken: string;
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
        salesforceUrl: this.properties.salesforceUrl,
        salesforceToken: this.properties.salesforceToken,
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
                PropertyPaneTextField('salesforceUrl', {
                  label: 'URL til endepunkt i Salesforce',
                }),
                PropertyPaneTextField('salesforceToken', {
                  label: 'Token (uten `Bearer`)',
                }),
              ]
            },
            {
              groupName: 'Oppdater token',
              groupFields: [
                PropertyPaneTextField('accessTokenUrl', {
                  label: 'Access token URL',
                }),
                PropertyPaneTextField('accessTokenUserName', {
                  label: 'Brukernavn',
                }),
                PropertyPaneTextField('accessTokenPassword', {
                  label: 'Passord',
                }),
                PropertyPaneTextField('accessTokenClientID', {
                  label: 'Client ID',
                }),
                PropertyPaneTextField('accessTokenSecret', {
                  label: 'Secret',
                }),
                PropertyPaneButton('accessToken', {
                  text: 'Get token',
                  onClick: this.getToken,
                }),
              ]
            }
          ]
        }
      ]
    };
  }

  protected getToken = async (event: React.FormEvent<HTMLFormElement>) => {
    const body = JSON.stringify({
      'grant_type': 'password',
      'client_id': this.properties.accessTokenClientID,
      'client_secret': this.properties.accessTokenSecret,
      'username': this.properties.accessTokenUserName,
      'password': this.properties.accessTokenPassword,
    });
    const headers: Headers = new Headers({
      'X-PrettyPrint': '1',
    });
    const httpClientOptions: IHttpClientOptions = {body, headers};
    try {
      const response: HttpClientResponse = await this.context.httpClient.post(
        this.properties.accessTokenUrl,
        HttpClient.configurations.v1,
        httpClientOptions,
      );
      await response.json();
      console.log(response);
    } catch (e) {
      console.error(e);
    } finally {
    }
  }

}
