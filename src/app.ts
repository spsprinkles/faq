import { FilterSlideout, Footer, ItemForm, Navigation } from "dattatable";
import { Components } from "gd-sprest-bs";
import { DataSource } from "./ds";
import Strings from "./strings";

/**
 * Main Application
 */
export class App {
    // Constructor
    constructor(el: HTMLElement) {
        // Set the list name
        ItemForm.ListName = Strings.Lists.FAQ;

        // Initialize the application
        DataSource.init().then(() => {
            // Render the dashboard
            this.render(el);
        });
    }

    // Renders the dashboard
    private render(el: HTMLElement) {
        // Render the navigation
        this.renderNavigation(el);

        // Render the accordion
        this.renderAccordion(el);

        // Render the footer
        new Footer({
            el,
            itemsEnd: [{
                text: Strings.Version
            }]
        });
    }

    // Renders the accordion
    private renderAccordion(el: HTMLElement) {
        // Parse the items
        let accordionItems: Array<Components.IAccordionItem> = [];
        for (let i = 0; i < DataSource.Items.length; i++) {
            var item = DataSource.Items[i];
            let category = (item.Category || "").toLowerCase().replace(/ /g, '-');

            // Add an accordion item
            accordionItems.push({
                showFl: i == 0,
                header: item["Title"] || "",
                className: category,
                content: item["Answer"] || ""
            });
        }

        // Render the accordion
        Components.Accordion({
            el,
            id: "accordion-list",
            items: accordionItems
        });
    }

    // Renders the navigation
    private renderNavigation(el: HTMLElement) {
        // Render the filters
        let filter = new FilterSlideout({
            filters: [{
                header: "By Category",
                items: DataSource.CategoryFilters,
                onFilter: (filter: string) => {
                    let className = filter ? filter.toLowerCase().replace(/ /g, "-") : null;

                    // Parse all accordion items
                    let items = el.querySelectorAll(".accordion-item");
                    for (let i = 0; i < items.length; i++) {
                        let elItem = items[i];

                        // Show the item
                        elItem.classList.remove("d-none");

                        // See if a class name exists
                        if (className) {
                            // See if this item doesn't matches
                            if (!elItem.classList.contains(className)) {
                                // Hide the item
                                elItem.classList.add("d-none");
                            }
                        }
                    }
                }
            }]
        });

        // Render the navigation
        new Navigation({
            el,
            title: Strings.ProjectName,
            onShowFilter: () => {
                // Show the filter
                filter.show();
            },
            onRendering: props => {
                // Update the navigation properties
                props.type = Components.NavbarTypes.Primary;
            },
            items: [
                {
                    className: "btn-outline-light",
                    text: "Submit a Question",
                    isButton: true,
                    onClick: () => {
                        // Create an item
                        ItemForm.create({
                            onCreateEditForm: props => {
                                // Update the fields to display
                                props.excludeFields = ["Answer", "Approved"];

                                // Return the properties
                                return props;
                            },
                            onFormButtonsRendering: buttons => {
                                // Update the create button
                                buttons[0].text = "Submit";

                                // Return the buttons
                                return buttons;
                            }
                        });
                    }
                }
            ]
        });
    }
}