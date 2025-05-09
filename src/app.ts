import { Dashboard } from "dattatable";
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
    private _ds: DataSource = null;

    // Constructor
    constructor(ds: DataSource, el: HTMLElement, title: string = Strings.ProjectName, allowMultipleFilters: boolean, showCategory: boolean) {
        this._ds = ds;

        // Render the dashboard
        this.render(el, title, allowMultipleFilters, showCategory);
    }

    // Renders the navigation
    private generateNavItems() {
        let itemsEnd: Components.INavbarItem[] = [];

        // Create the settings menu items
        if (Security.IsAdmin) {
            itemsEnd.push(
                {
                    className: "btn-icon btn-outline-light me-2 p-2 py-1",
                    text: "Settings",
                    iconSize: 22,
                    iconType: gearWideConnected,
                    isButton: true,
                    items: [
                        {
                            text: "Manage " + this._ds.FaqList.ListName + " list",
                            onClick: () => {
                                // Show the FAQ list in a new tab
                                window.open(this._ds.FaqList.ListUrl, "_blank");
                            }
                        }
                    ]
                }
            );
        }

        // Add Admin only items
        if (Security.IsAdmin) {
            // Add the app settings
            itemsEnd[0].items.unshift(
                {
                    text: "App Settings",
                    onClick: () => {
                        // Show the install modal
                        InstallationModal.show(true);
                    }
                }
            );

            // Add the default security groups
            itemsEnd[0].items.push(
                {
                    text: Security.AdminGroup.Title + " Group",
                    onClick: () => {
                        // Show the settings in a new tab
                        window.open(Strings.SourceUrl + "/_layouts/15/people.aspx?MembershipGroupId=" + Security.AdminGroup.Id);
                    }
                },
                {
                    text: Security.MemberGroup.Title + " Group",
                    onClick: () => {
                        // Show the settings in a new tab
                        window.open(Strings.SourceUrl + "/_layouts/15/people.aspx?MembershipGroupId=" + Security.MemberGroup.Id);
                    }
                },
                {
                    text: Security.VisitorGroup.Title + " Group",
                    onClick: () => {
                        // Show the settings in a new tab
                        window.open(Strings.SourceUrl + "/_layouts/15/people.aspx?MembershipGroupId=" + Security.VisitorGroup.Id);
                    }
                }
            );
        }

        // Return the nav items
        return itemsEnd;
    }

    // Returns the FAQ icon as an SVG element
    private getFaqIcon(height?, width?, className?) {
        // Set the default values
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
    private render(el: HTMLElement, title: string, allowMultipleFilters: boolean, showCategory: boolean) {
        // Render the accordion
        let dashboard = new Dashboard({
            el,
            accordion: {
                items: this._ds.FaqList.Items,
                bodyFields: ["Answer"],
                filterFields: ["Category"],
                paginationLimit: Strings.PaginationLimit,
                titleFields: showCategory ? ["Category", "Title"] : ["Title"],
                titleTemplate: showCategory ? `<span class="badge text-bg-primary me-1">{Category}</span>{Title}` : null
            },
            filters: {
                items: [{
                    header: "By Category",
                    items: this._ds.CategoryFilters,
                    multi: allowMultipleFilters
                }]
            },
            footer: {
                onRendering: props => {
                    // Update the properties
                    props.className = "footer p-0";
                },
                onRendered: (el) => {
                    el.querySelector("nav.footer").classList.remove("bg-light");
                    el.querySelector("nav.footer .container-fluid").classList.add("p-0");
                },
                itemsEnd: [{
                    className: "pe-none",
                    text: "v" + Strings.Version,
                    onRender: (el) => {
                        // Hide version footer in a modern page
                        Strings.IsClassic ? null : el.classList.add("d-none");
                    }
                }]
            },
            navigation: {
                itemsEnd: this.generateNavItems(),
                searchPlaceholder: "Search...",
                showFilter: false,
                showSearch: true,
                title,
                onRendering: props => {
                    // Update the navigation properties
                    props.className = "navbar-expand rounded-top";
                    props.type = Components.NavbarTypes.Primary;

                    // Add a logo to the navbar brand
                    let div = document.createElement("div");
                    let text = div.cloneNode() as HTMLDivElement;
                    div.className = "d-flex me-2";
                    text.className = "ms-2";
                    text.append(title);
                    div.appendChild(this.getFaqIcon());
                    div.appendChild(text);
                    props.brand = div;
                },
                onRendered: (el) => {
                    el.querySelector(".navbar-brand").classList.add("pe-none");
                }
            },
            subNavigation: {
                onRendering: props => {
                    props.className = "navbar-sub rounded-bottom";
                },
                onRendered: (el) => {
                    el.classList.remove("bg-light");
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
                                        // Show the new form
                                        this.showNewForm();
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
                                        dashboard.showFilter();
                                    }
                                },
                            });
                        }
                    }
                ]
            }
        });
    }

    // Displays the new form
    private showNewForm() {
        // Display the new form
        this._ds.FaqList.newForm({
            webUrl: Strings.SourceUrl,
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
                el.querySelector("h5").innerText = "Ask a Question";
            }
        });
    }
}