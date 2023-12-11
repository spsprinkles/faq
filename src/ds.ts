import { Components, List, Types } from "gd-sprest-bs";
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
    static get CategoryFilters(): Components.ICheckboxGroupItem[] {
        let uniqueCategories = {};
        let items: Components.ICheckboxGroupItem[] = [];

        // Parse the items
        for (let i = 0; i < this.Items.length; i++) {
            let item = this.Items[i];

            // Parse the selected categores
            let categories = item.Category ? item.Category.results : [];
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
            items.push({
                label: categoryNames[i],
                type: Components.CheckboxGroupTypes.Switch
            });
        }

        // Return the items
        return items;
    }

    // Initializes the application
    static init(): PromiseLike<any> {
        // Return a promise
        return new Promise((resolve, reject) => {
            return Promise.all([
                // Load the security
                Security.init(),
                // Load the data
                this.load()
            ]).then(resolve, reject);
        });
    }

    // Loads the list data
    private static _items: Array<IItem> = null;
    static get Items(): Array<IItem> { return this._items; }
    static load(): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Load the data
            List(Strings.Lists.FAQ).Items().query({
                Filter: "Approved eq 1",
                OrderBy: ["Title"],
                Top: 1000
            }).execute(
                // Success
                items => {
                    // Set the items
                    this._items = items.results as any;

                    // Resolve the request
                    resolve();
                },

                // Error
                reject
            );
        });
    }

    // Load Request
    private static _request: IItem = null;
    static loadRequest(itemId: number): PromiseLike<IItem> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Load the data
            List(Strings.Lists.FAQ).Items(itemId).query({
                Expand: ["Author", "Base"],
                Select: ["*", "Author/EMail", "Author/Title", "Base/Title"]
            }).execute(
                // Success
                item => {
                    // Set the requests
                    this._request = item as any;
                    // Resolve the request
                    resolve(this._request);
                },
                // Error
                reject
            );
        });
    }

    // Requests
    private static _requests: Array<IItem> = null;
    static get Requests(): Array<IItem> { return this._requests; }
    static loadRequests(): PromiseLike<Array<IItem>> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Load the data
            List(Strings.Lists.FAQ).Items().query({
                Expand: ["Author", "Base"],
                GetAllItems: true,
                OrderBy: ["Title"],
                Select: ["*", "Author/EMail", "Author/Title", "Base/Title"],
                Top: 5000
            }).execute(
                // Success
                items => {
                    // Set the requests
                    this._requests = items.results as any;

                    // Resolve the request
                    resolve(this._requests);
                },

                // Error
                reject
            );
        });
    }
}