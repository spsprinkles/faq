import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './FaqWebPart.module.scss';
import * as strings from 'FaqWebPartStrings';

// Reference the solution
import "../../../../dist/faq.min.js";
declare var DattaFAQ;

export interface IFaqWebPartProps {
  description: string;
}

export default class FaqWebPart extends BaseClientSideWebPart<IFaqWebPartProps> {

  public render(): void {
    // Render the webpart
    DattaFAQ.load(this.domElement, this.context.pageContext);
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
