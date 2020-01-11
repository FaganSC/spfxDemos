import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'MsgraphcreategroupWebPartStrings';
import Msgraphcreategroup from './components/Msgraphcreategroup';
import { IMsgraphcreategroupProps } from './components/IMsgraphcreategroupProps';
import { MSGraphClient } from '@microsoft/sp-http';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

export interface IMsgraphcreategroupWebPartProps {
  description: string;
}

export default class MsgraphcreategroupWebPart extends BaseClientSideWebPart <IMsgraphcreategroupWebPartProps> {
  public render(): void {
    this.context.msGraphClientFactory.getClient()
      .then((client: MSGraphClient): void => {
        const element: React.ReactElement<IMsgraphcreategroupProps> = React.createElement(
          Msgraphcreategroup,
          {
            graphClient: client
          }
        );
        ReactDom.render(element, this.domElement);
      });
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
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
