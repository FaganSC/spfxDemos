import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'MsGraphGroupsWebPartStrings';
import MsGraphGroups from './components/MsGraphGroups';
import { IMsGraphGroupsProps } from './components/IMsGraphGroupsProps';
import { MSGraphClient } from '@microsoft/sp-http';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

export interface IMsGraphGroupsWebPartProps {
  description: string;
}

export default class MsGraphGroupsWebPart extends BaseClientSideWebPart<IMsGraphGroupsWebPartProps> {
  public render(): void {
    this.context.msGraphClientFactory.getClient()
      .then((client: MSGraphClient): void => {
        const element: React.ReactElement<IMsGraphGroupsProps> = React.createElement(
          MsGraphGroups,
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
