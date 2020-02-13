import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';

import * as strings from 'MsGraphConnectionSampleWebPartStrings';
import MsGraphConnectionSample from './components/MsGraphConnectionSample';
import { IMsGraphConnectionSampleProps } from './components/IMsGraphConnectionSampleProps';

export interface IMsGraphConnectionSampleWebPartProps {
  description: string;
}

export default class MsGraphConnectionSampleWebPart extends BaseClientSideWebPart<IMsGraphConnectionSampleWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IMsGraphConnectionSampleProps > = React.createElement(
      MsGraphConnectionSample,
      {
        description: this.properties.description,
        context:this.context
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
