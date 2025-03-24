import { List } from "dattatable";
import { Components, Types } from "gd-sprest-bs";
import { Security } from "./security";
import Strings from "./strings";

// Item
export interface IItem extends Types.SP.ListItem {
    Alerted: boolean;
    Answer: string;
    Approved: boolean;
    Category?: { results: string[] };
}

/**
 * Data Source
 */
export class DataSource {
    // Category Filters
    private static _categoryFilters: Components.ICheckboxGroupItem[] = null;
    static get CategoryFilters(): Components.ICheckboxGroupItem[] { return this._categoryFilters; }

    // Initializes the application
    static init(viewName?: string): PromiseLike<any> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Initialize the security first
            // If the list doesn't exist, then it will require the Security class to be initialized
            Security.init().then(() => {
                // Load the data
                this.load(viewName).then(resolve, reject);
            });
        });
    }

    // Loads the list data
    private static _faqList: List<IItem> = null;
    static get FaqList(): List<IItem> { return this._faqList; }
    static load(viewName?: string): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Load the list
            this._faqList = new List({
                listName: Strings.Lists.FAQ,
                viewName: viewName || Strings.ViewName,
                webUrl: Strings.SourceUrl,
                itemQuery: {
                    Filter: "Approved eq 1",
                    OrderBy: ["Title"],
                    Top: 1000
                },
                onInitError: reject,
                onInitialized: () => {
                    let uniqueCategories = {};

                    // Clear the filter items
                    this._categoryFilters = [];

                    // Parse the items
                    for (let i = 0; i < this.FaqList.Items.length; i++) {
                        let item = this.FaqList.Items[i];

                        // Parse the selected categories
                        let categories = item.Category ? (typeof (item.Category) === "string" ? [item.Category] : item.Category.results) : [];
                        for (let j = 0; j < categories.length; j++) {
                            let category = categories[j];

                            // Set the category
                            if (category) { uniqueCategories[category] = true; }
                        }
                    }

                    // Parse the unique categories and get the unique names
                    let categoryNames = [];
                    for (let name in uniqueCategories) { categoryNames.push(name); }

                    // Parse the sorted items
                    categoryNames = categoryNames.sort();
                    for (let i = 0; i < categoryNames.length; i++) {
                        // Add an item
                        this._categoryFilters.push({
                            label: categoryNames[i],
                            type: Components.CheckboxGroupTypes.Switch
                        });
                    }

                    // Resolve the request
                    resolve();
                }
            });
        });
    }
}