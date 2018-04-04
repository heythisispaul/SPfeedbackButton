import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'FeedbackButtonWebPartStrings';
import FeedbackButton from './components/FeedbackButton';
import { IFeedbackButtonProps } from './components/IFeedbackButtonProps';

export interface IFeedbackButtonWebPartProps {
  buttonText: string;
}

export default class FeedbackButtonWebPart extends BaseClientSideWebPart<IFeedbackButtonWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IFeedbackButtonProps > = React.createElement(
      FeedbackButton,
      {
        buttonText: this.properties.buttonText
      }
    );

    ReactDom.render(element, this.domElement);
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
                PropertyPaneTextField('buttonText', {
                  label: "Button Text"
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
