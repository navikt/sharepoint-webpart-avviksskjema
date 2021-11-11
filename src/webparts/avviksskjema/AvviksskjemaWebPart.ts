import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'AvviksskjemaWebPartStrings';
import Avviksskjema from './components/Avviksskjema';
import { IAvviksskjemaProps } from './components/IAvviksskjemaProps';
import { sp } from '@pnp/sp';

export interface IAvviksskjemaWebPartProps {
  salesforceUrl: string;
  salesforceToken: string;
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
                  label: 'Token (uten `Bearer`)'
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
