import { InstallationRequired } from "dattatable";
import { Components } from "gd-sprest-bs";
import { Configuration } from "./cfg";
import { Security } from "./security";
import Strings from "./strings";

/**
 * Installation Modal
 */
export class InstallationModal {
    // Shows the modal
    static show(showFl: boolean = false) {
        // See if an installation is required
        InstallationRequired.requiresInstall({ cfg: Configuration }).then(installFl => {
            let customErrors: Components.IListGroupItem[] = [];

            // See if an install is required
            if (installFl || customErrors.length > 0 || showFl) {
                // Show the dialog
                InstallationRequired.showDialog({
                    errors: customErrors
                });
            } else {
                // Log
                console.error("[" + Strings.ProjectName + "] Error initializing the solution.");
            }
        });
    }
}