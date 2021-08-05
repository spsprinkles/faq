import { ContextInfo } from "gd-sprest-bs";
import { App } from "./app";
import { Configuration } from "./cfg";
import Strings from "./strings";

// Create the global variable for this solution
window[Strings.GlobalVariable] = {
    Configuration,
    render: (el, context?) => {
        // Set the page context if it exists
        if (context) { ContextInfo.setPageContext(context); }

        // Render the app
        new App(el);
    }
}

// Get the element and render the app if it is found
let elApp = document.querySelector("#" + Strings.AppElementId) as HTMLElement;
if (elApp) { new App(elApp); }