import { DisplayMode, Environment, Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneCheckbox, PropertyPaneLabel, PropertyPaneSlider, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import * as strings from 'FaqWebPartStrings';

export interface IFaqWebPartProps {
  enableLoading: boolean;
  listName: string;
  paginationLimit: number;
  title: string;
  viewName: string;
  webUrl: string;
}

// Reference the solution
import "main-lib";
declare const FaqApp: {
  description: string;
  enableLoading: boolean;
  listName: string;
  render: (props: {
    el: HTMLElement;
    context?: WebPartContext;
    displayMode?: DisplayMode;
    enableLoading?: boolean;
    envType?: number;
    paginationLimit?: number;
    title?: string;
    viewName?: string;
    listName?: string;
    sourceUrl?: string;
  }) => void;
  paginationLimit: number;
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
    if (typeof(this.properties.enableLoading) === "undefined") { this.properties.enableLoading = FaqApp.enableLoading; }
    if (!this.properties.listName) { this.properties.listName = FaqApp.listName; }
    if (!this.properties.paginationLimit) { this.properties.paginationLimit = FaqApp.paginationLimit; }
    if (!this.properties.title) { this.properties.title = FaqApp.title; }
    if (!this.properties.viewName) { this.properties.viewName = FaqApp.viewName; }
    if (!this.properties.webUrl) { this.properties.webUrl = this.context.pageContext.web.serverRelativeUrl; }

    // Render the application
    FaqApp.render({
      el: this.domElement,
      context: this.context,
      displayMode: this.displayMode,
      enableLoading: this.properties.enableLoading,
      paginationLimit: this.properties.paginationLimit,
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
                PropertyPaneSlider('paginationLimit', {
                  label: strings.PaginationLimitFieldLabel,
                  max: 30,
                  min: 1,
                  showValue: true
                }),
                PropertyPaneLabel('loadingLabel', {
                  text: strings.EnableLoadingFieldLabel
                }),
                PropertyPaneCheckbox("enableLoading", {
                  text: "Enabled"
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
