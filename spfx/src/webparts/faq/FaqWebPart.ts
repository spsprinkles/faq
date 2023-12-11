import { DisplayMode, Environment, Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneLabel, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import * as strings from 'FaqWebPartStrings';

export interface IFaqWebPartProps {
  description: string;
}

// Reference the solution
import "../../../../dist/faq.min.js";
declare const FaqApp: {
  description: string;
  render: (props: {
    el: HTMLElement;
    context?: WebPartContext;
    displayMode?: DisplayMode;
    envType?: number;
    sourceUrl?: string;
  }) => void;
  updateTheme: (currentTheme: Partial<IReadonlyTheme>) => void;
  version: string;
};

export default class FaqWebPart extends BaseClientSideWebPart<IFaqWebPartProps> {
  public render(): void {
    // Render the application
    FaqApp.render({
      el: this.domElement,
      context: this.context,
      displayMode: this.displayMode,
      envType: Environment.type
    });
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    // Update the theme
    FaqApp.updateTheme(currentTheme);
  }

  protected get dataVersion(): Version {
    return Version.parse(FaqApp.version);
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneLabel('version', {
                  text: "v" + FaqApp.version
                })
              ]
            }
          ],
          header: {
            description: FaqApp.description
          }
        }
      ]
    };
  }
}
