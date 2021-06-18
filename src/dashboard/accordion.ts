import { Components } from "gd-sprest-bs";
import { ItemForm } from "./itemForm";

/**
 * Accordion
 */
export class Accordion {
    private _el: HTMLElement = null;
    private _filterCategory = "";
    private _filterText = "";
    private _items: Array<any> = null;

    // Constructor
    constructor(el: HTMLElement, items) {
        // Save the input parameters
        this._el = el;
        this._items = items;

        // Render the filter
        this.renderFilter();

        // Render the accordion
        this.renderAccordion();
    }

    // Filters the accordion, based on the filter text value
    private filter() {
        let filterValue = this._filterText.toLowerCase();
        let showAllItems = filterValue == "";

        // Get the accordion items
        let elItems = this._el.querySelectorAll(".accordion-item");

        // Parse the items
        for (let i = 0; i < elItems.length; i++) {
            let elItem = elItems[i];
            let elBody = elItem.querySelector(".accordion-body") as HTMLDivElement;
            let elHeader = elItem.querySelector(".accordion-header") as HTMLDivElement;

            // See if we are showing all items
            if (showAllItems) {
                // Ensure this item is visible
                elItem.classList.remove("d-none");
            } else {
                // See if the text contains the value
                if ((elBody.innerText || "").toLowerCase().indexOf(filterValue) < 0 &&
                    (elHeader.innerText || "").toLowerCase().indexOf(filterValue) < 0) {
                    // Hide this item
                    elItem.classList.add("d-none");
                } else {
                    // Ensure this item is visible
                    elItem.classList.remove("d-none");
                }
            }
        }
    }

    // Get the categories
    private getCategories() {
        let categories = {};

        // Parse the items
        for (let i = 0; i < this._items.length; i++) {
            let item = this._items[i];

            // Set the category
            categories[item["Category"]] = true;
        }

        // Parse the keys
        let values = [];
        for (let key in categories) {
            // Add the unique value
            values.push(key);
        }

        // Return the categories
        return values.sort();
    }

    // Render the accordion
    private renderAccordion() {
        // Parse the items
        let accordionItems: Array<Components.IAccordionItem> = [];
        for (let i = 0; i < this._items.length; i++) {
            var item = this._items[i];

            // See if a category filter exists
            if (this._filterCategory) {
                // Ensure the item is from the same category
                if (item["Category"] != this._filterCategory) { continue; }
            }

            // Add an accordion item
            accordionItems.push({
                showFl: i == 0,
                header: item["Title"] || "",
                content: item["Answer"] || ""
            });
        }

        // Render the accordion
        Components.Accordion({
            el: this._el,
            id: "accordion-list",
            items: accordionItems
        });

        // See if a filter exists
        if (this._filterText) {
            // Apply the filter
            this.filter();
        }
    }

    // Render the filter
    private renderFilter() {
        let items: Array<Components.IDropdownItem> = [{
            text: "All Categories"
        }];

        // Get the unique categories
        let categories = this.getCategories();

        // Parse the categories
        for (let i = 0; i < categories.length; i++) {
            let category = categories[i];

            // Create a dropdown item
            items.push({
                text: category,
                value: category
            });
        }

        // Render the form
        Components.Form({
            el: this._el,
            className: "mb-2",
            rows: [
                {
                    columns: [
                        {
                            control: {
                                onControlRendered: ctrl => {
                                    // Render a create button
                                    Components.Button({
                                        el: ctrl.el,
                                        id: "create-request",
                                        text: "Create Request",
                                        type: Components.ButtonTypes.OutlinePrimary,
                                        onClick: () => {
                                            // Show the form
                                            ItemForm.create();
                                        }
                                    });
                                }
                            },
                        },
                        {
                            control: {
                                items,
                                type: Components.FormControlTypes.Dropdown,
                                onChange: (item) => {
                                    // Set the filter
                                    this._filterCategory = item.value || "";

                                    // Remove the accordion element
                                    let elAccordion = this._el.querySelector("#accordion-list");
                                    this._el.removeChild(elAccordion);

                                    // Render the accordion
                                    this.renderAccordion();
                                }
                            } as Components.IFormControlPropsDropdown
                        },
                        {
                            control: {
                                type: Components.FormControlTypes.TextField,
                                onChange: (value) => {
                                    // Setting the filter text
                                    this._filterText = value;

                                    // Filter the accordion items
                                    this.filter();
                                },
                                onControlRendered: (ctrl) => {
                                    let elInput = ctrl.textbox.el.querySelector(".form-control") as HTMLInputElement;
                                    if (elInput) {
                                        // Set the type
                                        elInput.type = "search";
                                        elInput.placeholder = "Search";
                                        elInput.setAttribute("aria-label", "Search");
                                    }
                                }
                            } as Components.IFormControlPropsTextField
                        }
                    ]
                }
            ]
        })
    }
}