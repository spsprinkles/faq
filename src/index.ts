import { LoadingDialog, waitForTheme } from "dattatable";
import { ContextInfo, ThemeManager } from "gd-sprest-bs";
import { App } from "./app";
import { Configuration } from "./cfg";
import { DataSource } from "./ds";
import { InstallationModal } from "./install";
import Strings, { setContext } from "./strings";

// Styling
import "./styles.scss";

// Properties
interface IProps {
    el: HTMLElement;
    context?: any;
    displayMode?: number;
    envType?: number;
    title?: string;
    listName?: string;
    viewName?: string;
    sourceUrl?: string;
}

// Create the global variable for this solution
const GlobalVariable = {
    App: null,
    Configuration,
    description: Strings.ProjectDescription,
    listName: Strings.Lists.FAQ,
    render: (props: IProps) => {
        // Show a loading dialog
        LoadingDialog.setHeader("Loading FAQ App");
        LoadingDialog.setBody("This may take some time based on the number of FAQ's...");
        LoadingDialog.show();

        // Set the page context if it exists
        if (props.context) {
            // Set the context
            setContext(props.context, props.envType, props.sourceUrl);

            // Update the configuration
            Configuration.setWebUrl(props.sourceUrl || ContextInfo.webServerRelativeUrl);

            // See if the list name is set
            if (props.listName) {
                // Update the configuration
                Strings.Lists.FAQ = props.listName;
                Configuration._configuration.ListCfg[0].ListInformation.Title = props.listName;
            }
        }

        // Update the ProjectName from SPFx title field
        props.title ? Strings.ProjectName = props.title : null;

        // Initialize the application
        DataSource.init(props.viewName).then(
            // Success
            () => {
                // Update the loading dialog
                LoadingDialog.setHeader("Loading Theme");

                // Wait for the theme to be loaded
                waitForTheme().then(() => {
                    // Create the application
                    GlobalVariable.App = new App(props.el);
                    
                    // Hide the loading dialog
                    LoadingDialog.hide();
                });
            },
            // Error
            () => {
                // Show the installation modal
                InstallationModal.show();

                // Hide the loading dialog
                LoadingDialog.hide();
            });
    },
    title: Strings.ProjectName,
    updateTheme: (themeInfo) => {
        // Set the theme
        ThemeManager.setCurrentTheme(themeInfo);
    },
    version: Strings.Version,
    viewName: Strings.ViewName
}
window[Strings.GlobalVariable] = GlobalVariable;

// Get the element and render the app if it is found
let elApp = document.querySelector("#" + Strings.AppElementId) as HTMLElement;
if (elApp) {
    // Remove the extra border spacing on the webpart
    let contentBox = document.querySelector("#contentBox table.ms-core-tableNoSpace");
    contentBox ? contentBox.classList.remove("ms-webpartPage-root") : null;
    
    // Render the application
    GlobalVariable.render({ el: elApp });
}