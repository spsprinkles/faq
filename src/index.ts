import { ContextInfo } from "gd-sprest-bs";
import { App } from "./app";
import { Configuration } from "./cfg";
import { DataSource } from "./ds";
import { InstallationModal } from "./install";
import Strings, { setContext } from "./strings";

// Create the global variable for this solution
const GlobalVariable = {
    Configuration,
    render: (el, context?, sourceUrl?: string) => {
        // Set the page context if it exists
        if (context) {
            // Set the context
            setContext(context, sourceUrl);

            // Update the configuration
            Configuration.setWebUrl(sourceUrl || ContextInfo.webServerRelativeUrl);
        }

        // Initialize the application
        DataSource.init().then(
            // Success
            () => {
                // Render the app
                new App(el);
            },
            // Error
            () => {
                // Show the installation modal
                InstallationModal.show();
            });
    }
}
window[Strings.GlobalVariable] = GlobalVariable;

// Get the element and render the app if it is found
let elApp = document.querySelector("#" + Strings.AppElementId) as HTMLElement;
if (elApp) {
    // Remove the extra border spacing on the webpart
    let contentBox = document.querySelector("#contentBox table.ms-core-tableNoSpace");
    contentBox ? contentBox.classList.remove("ms-webpartPage-root") : null;
    
    // Render the application
    GlobalVariable.render(elApp);
}