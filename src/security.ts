import { ListSecurity, ListSecurityDefaultGroups } from "dattatable";
import { ContextInfo, Helper, SPTypes, Types } from "gd-sprest-bs";
import Strings from "./strings";

/**
 * Security
 * Code related to the security groups the user belongs to.
 */
export class Security {
    private static _listSecurity: ListSecurity = null;

    // Admin
    private static _isAdmin: boolean = false;
    static get IsAdmin(): boolean { return this._isAdmin; }
    private static _adminGroup: Types.SP.GroupOData = null;
    static get AdminGroup(): Types.SP.GroupOData { return this._adminGroup; }

    // Members
    private static _memberGroup: Types.SP.GroupOData = null;
    static get MemberGroup(): Types.SP.GroupOData { return this._memberGroup; }

    // Visitors
    private static _visitorGroup: Types.SP.GroupOData = null;
    static get VisitorGroup(): Types.SP.GroupOData { return this._visitorGroup; }

    // Initializes the class
    static init(listName: string = Strings.Lists.FAQ): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            this._listSecurity = new ListSecurity({
                webUrl: Strings.SourceUrl,
                listItems: [
                    {
                        listName,
                        groupName: ListSecurityDefaultGroups.Owners,
                        permission: SPTypes.RoleType.Administrator
                    },
                    {
                        listName,
                        groupName: ListSecurityDefaultGroups.Members,
                        permission: SPTypes.RoleType.Contributor
                    },
                    {
                        listName,
                        groupName: ListSecurityDefaultGroups.Visitors,
                        permission: SPTypes.RoleType.Contributor
                    }
                ],
                onGroupCreated: group => {
                    // Set the group owner
                    Helper.setGroupOwner(group.Title, this._adminGroup.Title).then(() => {
                        // Log
                        console.log("[" + group.Title + " Group] The owner was updated successfully to " + this._adminGroup.Title + ".");
                    });
                },
                onGroupsLoaded: () => {
                    // Set the groups
                    this._adminGroup = this._listSecurity.getGroup(ListSecurityDefaultGroups.Owners);
                    this._memberGroup = this._listSecurity.getGroup(ListSecurityDefaultGroups.Members);
                    this._visitorGroup = this._listSecurity.getGroup(ListSecurityDefaultGroups.Visitors);

                    // Set the user flags
                    this._isAdmin = this._listSecurity.CurrentUser.IsSiteAdmin || this._listSecurity.isInGroup(ContextInfo.userId, ListSecurityDefaultGroups.Owners);

                    // Resolve the request
                    resolve();
                }
            });
        });
    }

    static hasPermissions(): PromiseLike<boolean> {
        // See if the user has permissions
        return this._listSecurity.checkUserPermissions();
    }

    // Displays the security group configuration
    static show(onComplete: () => void) {
        // Create the groups
        this._listSecurity.show(true, onComplete);
    }
}