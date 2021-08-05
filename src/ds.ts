import { Components, List, Types } from "gd-sprest-bs";
import Strings from "./strings";

// Item
export interface IItem extends Types.SP.ListItem {
    Answer: string;
    Approved: boolean;
    Category: string;
}

/**
 * Data Source
 */
export class DataSource {
    // Category Filters
    static get CategoryFilters(): Components.ICheckboxGroupItem[] {
        let categories = {};
        let items: Components.ICheckboxGroupItem[] = [];

        // Parse the items
        for (let i = 0; i < this.Items.length; i++) {
            // Set the category
            let category = this.Items[i].Category;
            if (category) { categories[category] = true; }
        }

        // Parse the categories and get the unique names
        let categoryNames = [];
        for (let name in categories) { categoryNames.push(name); }

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
    static init(): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Load the data
            this.load().then(() => {
                // Resolve the request
                resolve();
            }, reject);
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
}