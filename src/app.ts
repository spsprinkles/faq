import { FilterSlideout, Footer, ItemForm, Navigation } from "dattatable";
import { Components } from "gd-sprest-bs";
import { plus } from "gd-sprest-bs/build/icons/svgs/plus";
import { questionCircleFill } from "gd-sprest-bs/build/icons/svgs/questionCircleFill";
import { DataSource } from "./ds";
import strings from "./strings";
import Strings from "./strings";

/**
 * Main Application
 */
export class App {
    // Constructor
    constructor(el: HTMLElement) {
        // Set the list name
        ItemForm.ListName = Strings.Lists.FAQ;

        // Render the dashboard
        this.render(el);
    }

    // Renders the dashboard
    private render(el: HTMLElement) {
        // Render the navigation
        this.renderNavigation(el);
        
        // Render the sub-navigation
        this.renderSubNavigation(el);

        // Render the accordion
        this.renderAccordion(el);

        // Render the footer
        new Footer({
            el,
            onRendering: props => {
                // Update the properties
                props.className = "navbar-expand-sm";
            },
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
        // Render the navigation
        new Navigation({
            el,
            title: Strings.ProjectName,
            hideFilter: true,
            onRendering: props => {
                // Update the navigation properties
                props.className = "bg-sharepoint navbar-expand-sm rounded-top";
                props.type = Components.NavbarTypes.Primary;
                
                // Add a logo to the navbar brand
                let div = document.createElement("div");
                let text = div.cloneNode() as HTMLDivElement;
                div.className = "d-flex me-2";
                text.className = "ms-2";
                text.append(Strings.ProjectName);
                div.appendChild(questionCircleFill());
                div.appendChild(text);
                props.brand = div;
            },
            onRendered: (el) => {
                el.querySelector("a.navbar-brand").classList.add("pe-none");
            },
            onSearch: value => {
                // Remove the spaces from the value
                value = (value ? value.trim() : "").toLowerCase();

                // Parse all accordion items
                let items = el.querySelectorAll(".accordion-item");
                for (let i = 0; i < items.length; i++) {
                    let elItem = items[i] as HTMLElement;
                    let elContent = elItem.querySelector(".accordion-body") as HTMLElement;

                    // Show the item
                    elItem.classList.remove("d-none");

                    // See if a search value exists
                    if (value) {
                        // See if the item contains the search value
                        if (elItem.innerText.toLowerCase().indexOf(value) < 0 &&
                            elContent.innerText.toLowerCase().indexOf(value) < 0) {
                            // Hide the item
                            elItem.classList.add("d-none");
                        }
                    }
                }
            },
            onSearchRendered: (el) => {
                el.setAttribute("placeholder", "Search all FAQ's");
            }
        });
    }

    // Renders the sub-navigation
    private renderSubNavigation(el: HTMLElement) {
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

        // Render the sub-navigation
        new Navigation({
            el,
            title: "",
            hideSearch: true,
            onRendering: props => {
                // Update the navigation properties
                props.className = "navbar-expand-sm navbar-sub rounded-bottom";
                props.type = Components.NavbarTypes.Light;
            },
            // Move the filter icon in front of the last icon
            onRendered: (el) => {
                let filter = el.querySelector(".filter-icon");
                if (filter) {
                    let filterItem = document.createElement("li");
                    filterItem.className = "nav-item";
                    filterItem.appendChild(filter);
                    let ul = el.querySelector("#navbar_content ul:last-child");
                    let li = ul.querySelector("li:last-child");
                    ul.insertBefore(filterItem, li);
                }
            },
            onShowFilter: () => {
                // Show the filter
                filter.show();
            },
            itemsEnd: [
                {
                    text: "Ask a Question",
                    onRender: (el, item) => {
                        // Clear the existing button
                        el.innerHTML = "";
                        // Create a span to wrap the icon in
                        let span = document.createElement("span");
                        span.className = "bg-white d-inline-flex ms-2 rounded";
                        el.appendChild(span);

                        // Render a tooltip
                        Components.Tooltip({
                            el: span,
                            content: item.text,
                            btnProps: {
                                // Render the icon button
                                iconType: plus,
                                iconSize: 28,
                                type: Components.ButtonTypes.OutlineSecondary,
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
                            },
                        });
                    }
                }
            ]
        });
    }
}