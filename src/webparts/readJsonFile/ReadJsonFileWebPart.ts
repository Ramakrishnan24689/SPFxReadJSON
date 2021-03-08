import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { setup as pnpSetup } from "@pnp/common";
import * as strings from 'ReadJsonFileWebPartStrings';
import ReadJsonFile from './components/ReadJsonFile';
import { IReadJsonFileProps } from './components/IReadJsonFileProps';

export interface IReadJsonFileWebPartProps {
  description: string;
}

export default class ReadJsonFileWebPart extends BaseClientSideWebPart<IReadJsonFileWebPartProps> {
  protected onInit(): Promise<void> {

    return super.onInit().then(_ => {
  
      // other init code may be present
  
      pnpSetup({
        spfxContext: this.context
      });
    });
  }
  public render(): void {
    const element: React.ReactElement<IReadJsonFileProps> = React.createElement(
      ReadJsonFile,
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
