import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, Environment, EnvironmentType } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'ClientListWebPartStrings';
import ClientList from './components/ClientList';
import { IClientListProps } from './components/IClientListProps';

import { sp } from "@pnp/sp";

export interface IClientListWebPartProps {
  description: string;
}

export default class ClientListWebPart extends BaseClientSideWebPart<IClientListWebPartProps> {

  public onInit(): Promise<void> {
    if(Environment.type === EnvironmentType.Local){
      sp.setup({
        sp: {
          baseUrl: 'https://localhost:4323/'
        }
      });
    } else {
      sp.setup({
        spfxContext: this.context
      });
    }
    return Promise.resolve();
  }

  public render(): void {
    const element: React.ReactElement<IClientListProps > = React.createElement(
      ClientList,
      {
        description: this.properties.description
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
