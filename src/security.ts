import { ContextInfo, Helper, SPTypes, Types, Web } from "gd-sprest-bs";
import Strings from "./strings";

/**
 * Security
 * Code related to the security groups the user belongs to.
 */
export class Security {
    // Admin
    private static _isAdmin: boolean = false;
    static get IsAdmin(): boolean { return this._isAdmin; }
    private static _adminGroup: Types.SP.GroupOData = null;
    static get AdminGroup(): Types.SP.GroupOData { return this._adminGroup; }

    // FAQ Manager
    private static _isFAQMgr: boolean = false;
    static get IsFAQMgr(): boolean { return this._isFAQMgr; }
    private static _faqMgrGroup: Types.SP.GroupOData = null;
    static get FAQMgrGroup(): Types.SP.GroupOData { return this._faqMgrGroup; }

    // Manager Emails
    private static _managerEmails: Array<string> = null;
    static get ManagerEmails(): Array<string> { return this._managerEmails; }

    // Creates the security groups
    static create(): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Parse the groups to create
            Helper.Executor([Strings.SecurityGroups.FAQMgr], groupInfo => {
                // Return a promise
                return new Promise((resolve, reject) => {
                    let web = Web(Strings.SourceUrl);

                    // Get the group
                    web.SiteGroups().getByName(groupInfo.Name).execute(
                        // Exists
                        group => {
                            // Resolve the request
                            resolve(null);
                        },

                        // Doesn't exist
                        () => {
                            // Create the group
                            Web(Strings.SourceUrl).SiteGroups().add({
                                AllowMembersEditMembership: false,
                                Title: groupInfo.Name,
                                Description: groupInfo.Description,
                                OnlyAllowMembersViewMembership: false
                            }).execute(
                                // Successful
                                group => {
                                    // Resolve the request
                                    resolve(null);
                                },

                                // Error
                                () => {
                                    // The user is probably not an admin
                                    console.error("Error creating the security group.");

                                    // Reject the request
                                    reject();
                                }
                            );
                        }
                    );
                });
            }).then(() => {
                // Re-initialize this class
                this.init().then(() => {
                    // Configure the security groups
                    this.configure().then(() => {
                        // Resolve the request
                        resolve();
                    }, reject);
                }, reject);
            });
        });
    }

    // Configures the security groups
    private static configure(): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Complete the async methods
            Promise.all([
                // Reset the list permissions
                this.resetListPermissions(),
                // Get the owners group
                this.getOwnersGroup(),
                // Get the everyone user
                this.getEveryoneUser(),
                // Get the definitions
                this.getPermissionTypes()
            ]).then((requests) => {
                let everyoneUser = requests[2];
                let ownersGroup = requests[1];
                let permissions = requests[3];

                // Check if the FAQ Manager group owner is the associated owners group
                if (this.FAQMgrGroup.Owner.Id != ownersGroup.Id) {
                    Helper.setGroupOwner(this.FAQMgrGroup.Title, ownersGroup.Title).then(() => {
                        // Log
                        console.log("[" + this.FAQMgrGroup.Title + " Group] The owner was updated successfully to " + ownersGroup.Title + ".");
                    });
                }
                
                // Get the web for the lists
                let web = Web(Strings.SourceUrl);

                for (let key in Strings.Lists) {
                    // Get the list to update
                    let list = web.Lists(Strings.Lists[key]);

                    // Ensure the everyone user exists
                    if (everyoneUser) {
                        // Set the list permissions
                        list.RoleAssignments().addRoleAssignment(everyoneUser.Id, permissions[SPTypes.RoleType.Reader]).execute(() => {
                            // Log
                            console.log("[" + Strings.Lists[key] + " List] The everyone claim was added successfully.");
                        });
                    }

                    // Ensure the owners group exists
                    if (ownersGroup) {
                        // Set the list permissions
                        list.RoleAssignments().addRoleAssignment(ownersGroup.Id, permissions[SPTypes.RoleType.Administrator]).execute(() => {
                            // Log
                            console.log("[" + Strings.Lists[key] + " List] The owners group was added successfully.");
                        });
                    }

                    // Ensure the FAQ Manager group exists
                    if (this.FAQMgrGroup) {
                        // Set the list permissions
                        list.RoleAssignments().addRoleAssignment(this.FAQMgrGroup.Id, permissions[SPTypes.RoleType.Contributor]).execute(() => {
                            // Log
                            console.log("[" + Strings.Lists[key] + " List] The FAQ Manager permission was added successfully.");
                        });
                    }

                    // Wait for the requests to complete
                    list.done(() => {
                        // Log
                        console.log("[" + Strings.Lists[key] + " List] The permissions settings for this list are complete.");
                    });
                }

                // Wait for all the requests to complete
                web.done(() => {
                    // Log
                    console.log("[All Lists] All permissions settings are complete.");
                    // Resolve the request
                    resolve();
                });
            }, reject);
        });
    }

    // Gets the everyone user group
    private static getEveryoneUser(): PromiseLike<Types.SP.User> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Get the everyone user
            Web(Strings.SourceUrl).SiteUsers().query({
                Filter: "Title eq 'Everyone' or Title eq 'Everyone except external users'"
            }).execute(users => {
                // Resolve the request
                resolve(users.results[0] as any);
            }, reject);
        });
    }

    // Gets the owners user group
    private static getOwnersGroup(): PromiseLike<Types.SP.Group> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Get the owners group
            Web(Strings.SourceUrl).AssociatedOwnerGroup().execute(group => {
                // Resolve the request
                resolve(group);
            }, reject);
        });
    }

    // Gets the role definitions for the permission types
    private static getPermissionTypes(): PromiseLike<{ [name: number]: number }> {
        // Return a promise
        return new Promise(resolve => {
            // Get the definitions
            Web(Strings.SourceUrl).RoleDefinitions().execute(roleDefs => {
                let roles = {};

                // Parse the role definitions
                for (let i = 0; i < roleDefs.results.length; i++) {
                    let roleDef = roleDefs.results[i];

                    // Add the role by type
                    roles[roleDef.RoleTypeKind > 0 ? roleDef.RoleTypeKind : roleDef.Name] = roleDef.Id;
                }

                // Resolve the request
                resolve(roles);
            });
        });
    }

    // Initialization
    static init(): PromiseLike<any> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Complete the async methods
            Promise.all([
                // Load the owners group
                this.loadOwnersGroup(),
                // Load the FAQ Manager group
                this.loadFAQMgrGroup()
            ]).then(resolve, reject);
        });
    }

    // Load the FAQ Manager group
    private static loadFAQMgrGroup(): PromiseLike<void> {
        // Return a Promise
        return new Promise((resolve, reject) => {
            // Get the FAQ Manager group
            Web(Strings.SourceUrl).SiteGroups().getByName(Strings.SecurityGroups.FAQMgr.Name).query({
                Expand: ["Owner", "Users"],
                Select: ["*", "Owner/Id"]
            }).execute(
                group => {
                    // Set the group
                    this._faqMgrGroup = group;

                    // Clear the manager emails
                    this._managerEmails = [];

                    // Parse the users
                    for (let i = 0; i < group.Users.results.length; i++) {
                        // Append the email
                        this._managerEmails.push(group.Users.results[i].Email);
                    }
                    
                    // Parse the users
                    for (let i = 0; i < group.Users.results.length; i++) {
                        // See if the current user is in this group
                        if (group.Users.results[i].Id == ContextInfo.userId) {
                            // Set the flag and break from the loop
                            this._isFAQMgr = true;
                            break;
                        }
                    }

                    // Resolve the request
                    resolve();
                },

                // Group doesn't exist
                () => {
                    // Reject the request
                    reject();
                }
            )
        });
    }

    // Load the owner's group
    private static loadOwnersGroup(): PromiseLike<void> {
        // Return a Promise
        return new Promise((resolve) => {
            // Default the flag
            this._isAdmin = ContextInfo.isSiteAdmin;

            // Get the owners group
            Web(Strings.SourceUrl).AssociatedOwnerGroup().query({ Expand: ["Users"] }).execute(
                group => {
                    // Set the group
                    this._adminGroup = group;

                    // Parse the users
                    for (let i = 0; i < group.Users.results.length; i++) {
                        // See if the current user is in this group
                        if (group.Users.results[i].Id == ContextInfo.userId) {
                            // Set the flag and break from the loop
                            this._isAdmin = true;
                            break;
                        }
                    }

                    // Resolve the request
                    resolve();
                },

                // No access to the group
                () => {
                    // Resolve the request
                    resolve();
                }
            )
        });
    }

    // Clears the security groups for a list
    private static resetListPermissions(): PromiseLike<void> {
        // Return a promise
        return new Promise(resolve => {
            Helper.Executor([Strings.Lists.FAQ], listName => {
                // Return a promise
                return new Promise(resolve => {
                    // Get the list
                    let list = Web(Strings.SourceUrl).Lists(listName);

                    // Reset the permissions
                    list.resetRoleInheritance().execute();

                    // Clear the permissions
                    list.breakRoleInheritance(false, true).execute(true);

                    // Wait for the requests to complete
                    list.done(resolve);
                });
            }).then(resolve);
        });
    }
}