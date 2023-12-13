import { DisplayMode, Environment, Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneLabel, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import * as strings from 'FaqWebPartStrings';

export interface IFaqWebPartProps {
  title: string;
  listName: string;
  viewName: string;
  webUrl: string;
}

// Reference the solution
import "../../../../dist/faq.min.js";
declare const FaqApp: {
  description: string;
  listName: string;
  render: (props: {
    el: HTMLElement;
    context?: WebPartContext;
    displayMode?: DisplayMode;
    envType?: number;
    title?: string;
    viewName?: string;
    listName?: string;
    sourceUrl?: string;
  }) => void;
  title: string;
  updateTheme: (currentTheme: Partial<IReadonlyTheme>) => void;
  version: string;
  viewName: string;
};

export default class FaqWebPart extends BaseClientSideWebPart<IFaqWebPartProps> {
  private _hasRendered: boolean = false;

  public render(): void {
    // See if have rendered the solution
    if (this._hasRendered) {
      // Clear the element
      while (this.domElement.firstChild) { this.domElement.removeChild(this.domElement.firstChild); }
    }

    // Set the default property values
    if (!this.properties.listName) { this.properties.listName = FaqApp.listName; }
    if (!this.properties.title) { this.properties.title = FaqApp.title; }
    if (!this.properties.viewName) { this.properties.viewName = FaqApp.viewName; }
    if (!this.properties.webUrl) { this.properties.webUrl = this.context.pageContext.web.serverRelativeUrl; }

    // Render the application
    FaqApp.render({
      el: this.domElement,
      context: this.context,
      displayMode: this.displayMode,
      envType: Environment.type,
      title: this.properties.title,
      viewName: this.properties.viewName,
      listName: this.properties.listName,
      sourceUrl: this.properties.webUrl
    });

    // Set the flag
    this._hasRendered = true;
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

  protected get disableReactivePropertyChanges(): boolean { return true; }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupFields: [
                PropertyPaneTextField('title', {
                  label: strings.TitleFieldLabel,
                  description: strings.TitleFieldDescription
                }),
                PropertyPaneTextField('webUrl', {
                  label: strings.WebUrlFieldLabel,
                  description: strings.WebUrlFieldDescription
                }),
                PropertyPaneTextField('listName', {
                  label: strings.ListNameFieldLabel,
                  description: strings.ListNameFieldDescription
                }),
                PropertyPaneTextField('viewName', {
                  label: strings.ViewNameFieldLabel,
                  description: strings.ViewNameFieldDescription
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
