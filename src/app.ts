import { FilterSlideout, Footer, ItemForm, Navigation } from "dattatable";
import { Components } from "gd-sprest-bs";
import { filterSquare } from "gd-sprest-bs/build/icons/svgs/filterSquare";
import { link45deg } from "gd-sprest-bs/build/icons/svgs/link45deg";
import { questionCircleFill } from "gd-sprest-bs/build/icons/svgs/questionCircleFill";
import { questionSquare } from "gd-sprest-bs/build/icons/svgs/questionSquare";
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

        // Render the dashboard
        this.render(el);
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
            onRendering: props => {
                // Update the properties
                props.className = "footer p-0";
            },
            onRendered: (el) => {
                el.querySelector("nav.footer").classList.remove("bg-light");
            },
            itemsEnd: [{
                className: "pe-none",
                text: "v" + Strings.Version
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
            hideFilter: true,
            title: Strings.ProjectName,
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

        // Render the sub-navigation
        let subNav = Components.Navbar({
            el,
            className: "navbar-sub rounded-bottom",
            type: Components.NavbarTypes.Light,
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
                                className: "p-1 pe-2",
                                iconClassName: "me-1",
                                iconType: questionSquare,
                                iconSize: 24,
                                text: "Questions",
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
                                        },
                                        onSetHeader: el => {
                                            el.querySelector("h5").innerText = item.text;
                                        }
                                    });
                                }
                            },
                        });
                    }
                },
                {
                    text: "Filter the FAQ's",
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
                                className: "p-1 pe-2",
                                iconClassName: "me-1",
                                iconType: filterSquare,
                                iconSize: 24,
                                text: "Filters",
                                type: Components.ButtonTypes.OutlineSecondary,
                                onClick: () => {
                                    // Show the filter
                                    filter.show();
                                }
                            },
                        });
                    }
                },
                {
                    // Create Dropdown
                    text: "Links",
                    isButton: true,
                    className: "btn btn-sm btn-outline-secondary ms-2",
                    classNameItem: "bg-white",
                    iconSize: 22,
                    iconType: link45deg,
                    items: [
                        {
                            text: Strings.Links.Support.text,
                            onClick: () => {
                                window.location.href = Strings.Links.Support.href;
                            }
                        },
                        {
                            text: Strings.Links.GetSupport.text,
                            onClick: () => {
                                window.location.href = Strings.Links.GetSupport.href;
                            }
                        }
                    ]
                }
            ]
        });
    }
}