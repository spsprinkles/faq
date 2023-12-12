import { FilterSlideout, Footer, ItemForm, Navigation } from "dattatable";
import { Components, Utility, ContextInfo } from "gd-sprest-bs";
import { filterSquare } from "gd-sprest-bs/build/icons/svgs/filterSquare";
import { gearWideConnected } from "gd-sprest-bs/build/icons/svgs/gearWideConnected";
import { questionCircleFill } from "gd-sprest-bs/build/icons/svgs/questionCircleFill";
import { questionSquare } from "gd-sprest-bs/build/icons/svgs/questionSquare";
import { DataSource } from "./ds";
import { InstallationModal } from "./install";
import { Security } from "./security";
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
            onRendering: props => {
                // Update the properties
                props.className = "footer p-0";
            },
            onRendered: (el) => {
                el.querySelector("nav.footer").classList.remove("bg-light");
                el.querySelector("nav.footer .container-fluid").classList.add("p-0");
            },
            itemsEnd: [{
                className: "pe-none text-body",
                text: "v" + Strings.Version,
                onRender: (el) => {
                    // Hide version footer in a modern page
                    Strings.IsClassic ? null : el.classList.add("d-none");
                }
            }]
        });
    }

    // Renders the accordion
    private renderAccordion(el: HTMLElement) {
        // Parse the items
        let accordionItems: Array<Components.IAccordionItem> = [];
        for (let i = 0; i < DataSource.Items.length; i++) {
            let item = DataSource.Items[i];
            let itemClassNames = [];

            // Parse the selected categories
            let categories = item.Category ? item.Category.results : [];
            for (let j = 0; j < categories.length; j++) {
                let category = categories[j].toLowerCase().replace(/ /g, '-');

                // Set the category
                if (category) { itemClassNames.push(category); }
            }

            // Add an accordion item
            accordionItems.push({
                showFl: i == 0,
                header: item["Title"] || "",
                className: itemClassNames.join(" "),
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

        // Create the settings menu items
        let itemsEnd: Components.INavbarItem[] = [];
        if (Security.IsAdmin || Security.IsFAQMgr) {
            itemsEnd.push(
                {
                    className: "btn-icon btn-outline-light me-2 p-2 py-1",
                    text: "Settings",
                    iconSize: 22,
                    iconType: gearWideConnected,
                    isButton: true,
                    items: [
                        {
                            text: "Manage FAQs",
                            onClick: () => {
                                // Show the FAQ list in a new tab
                                window.open(Strings.SourceUrl + "/Lists/" + Strings.Lists.FAQ, "_blank");
                            }
                        }
                    ]
                }
            );
        }

        // Add Admin only items
        if (Security.IsAdmin) {
            itemsEnd[0].items.unshift(
                {
                    text: "App Settings",
                    onClick: () => {
                        // Show the install modal
                        InstallationModal.show(true);
                    }
                }
            );
            itemsEnd[0].items.push(
                {
                    text: "Manage Security",
                    onClick: () => {
                        // Show the settings in a new tab
                        window.open(Strings.SourceUrl + "/_layouts/15/people.aspx?MembershipGroupId=" + Security.FAQMgrGroup.Id);
                    }
                }
            );
        }

        // Render the navigation
        new Navigation({
            el,
            hideFilter: true,
            title: Strings.ProjectName,
            onRendering: props => {
                // Update the navigation properties
                props.className = "navbar-expand rounded-top";
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
            },
            itemsEnd
        });

        // Render the sub-navigation
        Components.Navbar({
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
                            placement: Components.TooltipPlacements.Left,
                            btnProps: {
                                // Render the icon button
                                className: "p-1 pe-2",
                                iconClassName: "me-1",
                                iconType: questionSquare,
                                iconSize: 24,
                                isSmall: true,
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
                                        onSave: (props) => {
                                            let categories = props.Category.results;
                                            // See if item is created
                                            Utility().sendEmail({
                                                To: Security.ManagerEmails,
                                                Subject: "New FAQ Request",
                                                Body: [
                                                    "<p>FAQ Managers,</p>",
                                                    "\t<p>" + ContextInfo.userDisplayName + " has created a new FAQ request.<br />",
                                                    "<b>Category(ies):</b> " + categories + "<br />",
                                                    "<b>Question:</b> " + props.Title + "<br />",
                                                    "<a href=\"" + ContextInfo.webAbsoluteUrl + "/Lists/" + Strings.Lists.FAQ + "\">Click here</a> to view the request.</p>",
                                                    "<p>Thank you,</p>",
                                                    "<p>FAQ Managers</p>"
                                                ].join('\n<br/>\n')
                                            }).execute();
                                            // Return the properties
                                            return props;
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
                            placement: Components.TooltipPlacements.Right,
                            btnProps: {
                                // Render the icon button
                                className: "p-1 pe-2",
                                iconClassName: "me-1",
                                iconType: filterSquare,
                                iconSize: 24,
                                isSmall: true,
                                text: "Filters",
                                type: Components.ButtonTypes.OutlineSecondary,
                                onClick: () => {
                                    // Show the filter
                                    filter.show();
                                }
                            },
                        });
                    }
                }
            ],
            onRendered: (el) => {
                el.classList.remove("bg-light");
            }
        });
    }
}