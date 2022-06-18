import { InstallationRequired } from "dattatable";
import { ContextInfo } from "gd-sprest-bs";
import { App } from "./app";
import { Configuration } from "./cfg";
import { DataSource } from "./ds";
import Strings from "./strings";

// Styling
import "./styles.scss";

// Create the global variable for this solution
const GlobalVariable = {
    Configuration,
    render: (el, context?) => {
        // Set the page context if it exists
        if (context) { ContextInfo.setPageContext(context); }

        // Initialize the application
        DataSource.init().then(
            // Success
            () => {
                // Render the app
                new App(el);
            },
            // Error
            () => {
                // Determine if an install is required
                InstallationRequired.requiresInstall(Configuration).then(requiresInstall => {
                    // See if an install is required
                    if (requiresInstall) {
                        // Show the installation required modal
                        InstallationRequired.showDialog();
                    }
                });
            });
    }
}
window[Strings.GlobalVariable] = GlobalVariable;

// Get the element and render the app if it is found
let elApp = document.querySelector("#" + Strings.AppElementId) as HTMLElement;
if (elApp) {
    // Render the application
    GlobalVariable.render(elApp);
}