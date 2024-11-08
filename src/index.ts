import { LoadingDialog, waitForTheme } from "dattatable";
import { Components, ContextInfo, ThemeManager } from "gd-sprest-bs";
import { App } from "./app";
import { Configuration } from "./cfg";
import { DataSource } from "./ds";
import { InstallationModal } from "./install";
import { Security } from "./security";
import Strings, { setContext } from "./strings";

// Styling
import "./styles.scss";

// Properties
interface IProps {
    el: HTMLElement;
    context?: any;
    displayMode?: number;
    enableLoading?: boolean;
    envType?: number;
    listName?: string;
    paginationLimit?: number;
    sourceUrl?: string;
    title?: string;
    viewName?: string;
}

// Create the global variable for this solution
const GlobalVariable = {
    App: null,
    Configuration,
    description: Strings.ProjectDescription,
    enableLoading: Strings.EnableLoading,
    listName: Strings.Lists.FAQ,
    paginationLimit: Strings.PaginationLimit,
    render: (props: IProps) => {
        // Set the EnableLoading value from SPFx settings
        (typeof (props.enableLoading) === "undefined") ? null : Strings.EnableLoading = props.enableLoading;

        // Show a loading dialog
        LoadingDialog.setHeader("Loading FAQ App");
        LoadingDialog.setBody("This may take some time based on the number of FAQ's...");
        Strings.EnableLoading ? LoadingDialog.show() : null;

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

        // Update the PaginationLimit from SPFx value
        props.paginationLimit ? Strings.PaginationLimit = props.paginationLimit : null;

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
                    Strings.EnableLoading ? LoadingDialog.hide() : null;
                });
            },
            // Error
            () => {
                // See if the user has the correct permissions
                Security.hasPermissions().then(hasPermissions => {
                    // See if the user has permissions
                    if (hasPermissions) {
                        // Show the installation modal
                        InstallationModal.show();
                    } else {
                        // Render an alert
                        Components.Alert({
                            header: "Permissions Issue",
                            content: "You do not have read access to the FAQ list. Please contact your admin to grant access for your account.",
                            type: Components.AlertTypes.Danger
                        });
                    }

                    // Hide the loading dialog
                    Strings.EnableLoading ? LoadingDialog.hide() : null;
                });
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