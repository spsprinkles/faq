import { DataSource } from "../ds";
import { Accordion } from "./accordion";
import { Footer } from "./footer";

// Styling
import "./styles.scss";

/**
 * Dashboard
 */
export class Dashboard {
    private _el: HTMLElement = null;

    // Constructor
    constructor(el: HTMLElement) {
        // Update the properties
        this._el = el;

        // Render the dashboard
        this.render();
    }

    // Renders the component
    private render() {
        // Render the template
        this._el.innerHTML = `<div id="faq"></div><div id="footer"></div>`;

        // Load the data
        DataSource.load().then(items => {
            // Render the Accordion
            new Accordion(this._el.querySelector("#faq"), items);
        });

        // Render the footer
        new Footer(this._el.querySelector("#footer"));
    }
}