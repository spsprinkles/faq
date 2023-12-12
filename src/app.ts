import { FilterSlideout, Footer, ItemForm, Navigation } from "dattatable";
import { Components, Utility, ContextInfo } from "gd-sprest-bs";
import { filterSquare } from "gd-sprest-bs/build/icons/svgs/filterSquare";
import { gearWideConnected } from "gd-sprest-bs/build/icons/svgs/gearWideConnected";
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

        // Render the dashboard
        this.render(el);
    }

    // Returns the FAQ icon as an SVG element
    private getFaqIcon(height?, width?, className?) {
        if (height === void 0) { height = 30; }
        if (width === void 0) { width = 36; }
        // Get the icon element
        let elDiv = document.createElement("div");
        elDiv.innerHTML = "<svg xmlns='http://www.w3.org/2000/svg' fill='none' viewBox='0 0 14.16 18.128572'><path d='m 3.2561612,3.3927144 c 0.434,-0.232 0.901,-0.306 1.356,-0.306 0.526,0 1.138,0.173 1.632,0.577 0.517,0.424 0.868,1.074 0.868,1.922 0,0.975 -0.689,1.504 -1.077,1.802 l -0.085,0.066 c -0.424,0.333 -0.588,0.511 -0.588,0.882 a 0.75,0.75 0 0 1 -1.5,0 c 0,-1.134 0.711,-1.708 1.162,-2.062 0.513,-0.403 0.588,-0.493 0.588,-0.688 0,-0.397 -0.149,-0.622 -0.32,-0.761 a 1.115,1.115 0 0 0 -0.68,-0.239 c -0.295,0 -0.498,0.049 -0.65,0.13 -0.143,0.076 -0.294,0.21 -0.44,0.48 a 0.75060401,0.75060401 0 1 1 -1.3200002,-0.715 c 0.264,-0.486 0.612,-0.853 1.0540002,-1.089 z m 1.356,8.6929996 a 1.0000001,1.0000001 0 1 0 0,-2 1.0000001,1.0000001 0 0 0 0,2 z'/><path d='m 4.6121612,0.08571439 a 7.5,7.5 0 0 0 -6.797,10.67299961 l -0.725,2.842 a 1.25,1.25 0 0 0 1.504,1.524 c 0.74999994,-0.18 1.90299994,-0.457 2.9299998,-0.702 A 7.5,7.5 0 1 0 4.6121612,0.08571439 Z m -6,7.50000001 A 6,6 0 1 1 1.942161,12.960714 l -0.243,-0.121 -0.265,0.063 -2.7879998,0.667 c 0.2,-0.78 0.46199994,-1.812 0.68999994,-2.708 l 0.07,-0.276 -0.13,-0.253 A 5.971,5.971 0 0 1 -1.3878388,7.5857144 Z'/><path d='m 9.6121612,18.085714 c -1.97,0 -3.761,-0.759 -5.1,-2 h 0.1 c 0.718,0 1.415,-0.089 2.081,-0.257 0.864,0.482 1.86,0.757 2.92,0.757 0.9599998,0 1.8659998,-0.225 2.6689998,-0.625 l 0.243,-0.121 0.265,0.063 c 0.921,0.22 1.965,0.445 2.74,0.61 -0.176,-0.751 -0.415,-1.756 -0.642,-2.651 l -0.07,-0.276 0.13,-0.253 a 5.971,5.971 0 0 0 0.664,-2.747 5.995,5.995 0 0 0 -2.747,-5.0419996 8.443,8.443 0 0 0 -0.8,-2.047 7.503,7.503 0 0 1 4.344,10.2629996 c 0.253,1.008 0.51,2.1 0.672,2.803 a 1.244,1.244 0 0 1 -1.468,1.5 c -0.727,-0.152 -1.87,-0.396 -2.913,-0.64 a 7.476,7.476 0 0 1 -3.0879998,0.663 z'/></svg>";
        let icon = elDiv.firstChild as SVGImageElement;
        if (icon) {
            // See if a class name exists
            if (className) {
                // Parse the class names
                let classNames = className.split(' ');
                for (var i = 0; i < classNames.length; i++) {
                    // Add the class name
                    icon.classList.add(classNames[i]);
                }
            } else {
                icon.classList.add("icon-svg");
            }
            // Set the height/width
            height ? icon.setAttribute("height", (height).toString()) : null;
            width ? icon.setAttribute("width", (width).toString()) : null;
            // Hide the icon as non-interactive content from the accessibility API
            icon.setAttribute("aria-hidden", "true");
            // Update the styling
            icon.style.pointerEvents = "none";
            // Support for IE
            icon.setAttribute("focusable", "false");
        }
        // Return the icon
        return icon;
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
                div.appendChild(this.getFaqIcon());
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