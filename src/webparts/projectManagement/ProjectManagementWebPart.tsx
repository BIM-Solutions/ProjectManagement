import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import * as strings from 'ProjectManagementWebPartStrings';
import { spfi } from "@pnp/sp";
import { SPFx } from "@pnp/sp/presets/all";
import { ListService } from './services/ListService';
import { SPProvider } from './components/common/SPContext';
import LandingPage from './components/LandingPage';
import { PropertyPaneButton, PropertyPaneButtonType } from '@microsoft/sp-property-pane';
import { LoadingProvider } from './services/LoadingContext';
import {
  FluentProvider,
  webLightTheme
} from "@fluentui/react-components";



export interface IProjectManagementWebPartProps {
  description: string;
}

export default class ProjectManagementWebPart extends BaseClientSideWebPart<IProjectManagementWebPartProps> {

  public _isDarkTheme: boolean = false;
  public _environmentMessage: string = '';

  /**
   * Lifecycle method that is called when the web part is initialized.
   * It ensures that all the required lists and their fields are provisioned on the site.
   * It also populates the _environmentMessage property with a message that is suitable for display
   * in the UI, indicating the current environment (e.g. local, SharePoint, Teams, etc.).
   */
  public async onInit(): Promise<void> {
    const sp = spfi().using(SPFx(this.context));
    const listService = new ListService(sp);
    const firstListAvailable = await listService.ensureFirstListSchema();
    if (!firstListAvailable) {
      console.warn("First required list is missing.");
      // Optionally store this in a global flag or react state via a context or callback to LandingPage
    }

    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }

  /**
   * Lifecycle method that is called when the web part is rendered.
   * It renders a React component that wraps the LandingPage component
   * in an SPProvider component, passing the WebPartContext to it.
   * This allows the LandingPage component to access the SP context and
   * use it to query the SharePoint REST API.
   */
  public render(): void {
    const element = (
      <FluentProvider theme={webLightTheme}>
        <SPProvider context={this.context}>
          <LoadingProvider>
            <LandingPage context={this.context} />
          </LoadingProvider>
        </SPProvider>
      </FluentProvider>
    );

    ReactDom.render(element, this.domElement);
  }

  /**
   * Determines the environment in which the web part is running and returns a corresponding message.
   * If the web part is running within Microsoft Teams, it retrieves the app context to distinguish between
   * Office, Outlook, and Teams environments and returns the appropriate message based on whether the web part
   * is served from localhost or not. If not in Teams, it defaults to checking if the environment is SharePoint.
   * The returned message is suitable for display in the UI to inform users about the current environment.
   *
   * @returns A promise that resolves to a string message indicating the current environment.
   */

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) {
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams':
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }
          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  /**
   * Lifecycle method that is called when the theme or style of the web part changes.
   * It updates the component's _isDarkTheme property and sets the semantic colors
   * for the web part's DOM element, if the currentTheme is defined.
   * @param currentTheme The current theme, if any.
   */
  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) return;

    this._isDarkTheme = !!currentTheme.isInverted;
    const semanticColors = currentTheme.semanticColors;
    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }
  }

/**
 * Lifecycle method that is called when the web part is disposed.
 * It unmounts the React component from the DOM element to clean up
 * resources and avoid memory leaks when the web part is removed from the page.
 */

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  /**
   * Gets the data version for the web part.
   * @returns The data version in the format of a Version object.
   * The Version object is used to determine whether the component
   * needs to be re-rendered when the data version changes.
   * The default implementation returns a Version object with
   * a fixed version number.
   */
  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  /**
   * Gets the configuration for the property pane.
   * @returns The property pane configuration.
   * The configuration is an object that defines the layout and fields of the property pane.
   * The property pane is used to configure the web part in the SharePoint page editor.
   * The configuration is a set of pages, each with a header and groups of fields.
   * The first page is the default page and is used to display the property pane.
   * The header of the page has a description field that displays the description of the web part.
   * The groups of fields are used to group related fields together.
   * The first group is the basic group, which contains a text field for the description of the web part.
   */
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
                }),
                PropertyPaneButton('verifyLists', {
                  text: "Check & Create Lists",
                  buttonType: PropertyPaneButtonType.Primary,
                  onClick: async () => {
                    const sp = spfi().using(SPFx(this.context));
                    const listService = new ListService(sp);

                    // get loading setter from DOM
                    interface CustomWindow extends Window {
                      __setListLoading?: (isLoading: boolean) => void;
                    }
                    const loadingSetter = (window as CustomWindow).__setListLoading;
                    if (loadingSetter) loadingSetter(true);

                    await listService.ensureListSchema();

                    if (loadingSetter) loadingSetter(false);
                    alert("List check and creation complete.");
                  }
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
