import { List, Types } from "gd-sprest-bs";
import Strings from "./strings";

// Item
export interface IItem extends Types.SP.ListItem { }

/**
 * Data Source
 */
export class DataSource {
    // Loads the list data
    static load(): PromiseLike<Array<IItem>> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Load the data
            List(Strings.Lists.FAQ).Items().query({
                Filter: "Approved eq 1",
                OrderBy: ["Title"],
                Top: 1000
            }).execute(
                // Success
                items => { resolve(items.results as any); },
                // Error
                reject
            );
        });
    }
}