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
    private static _adminGroup: Types.SP.Group = null;
    static get AdminGroup(): Types.SP.Group { return this._adminGroup; }

    // FAQ Manager
    private static _isFAQMgr: boolean = false;
    static get IsFAQMgr(): boolean { return this._isFAQMgr; }
    private static _faqMgrGroup: Types.SP.Group = null;
    static get FAQMgrGroup(): Types.SP.Group { return this._faqMgrGroup; }
    private static _faqMgrGroupInfo: Types.SP.GroupCreationInformation = {
        AllowMembersEditMembership: false,
        Description: "The FAQ managers.",
        OnlyAllowMembersViewMembership: false,
        Title: "FAQ Managers"
    };

    // Visitors
    private static _visitorGroup: Types.SP.Group = null;
    static get VisitorGroup(): Types.SP.Group { return this._visitorGroup; }

    // Manager Emails
    private static _managerEmails: Array<string> = null;
    static get ManagerEmails(): Array<string> { return this._managerEmails; }

    // Security Group Url
    private static _securityGroupUrl = ContextInfo.webServerRelativeUrl + "/_layouts/15/people.aspx?MembershipGroupId=";
    static get SecurityGroupUrl(): string { return this._securityGroupUrl };

    // Initializes the class
    static init(): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            this._listSecurity = new ListSecurity({
                groups: [this._faqMgrGroupInfo],
                webUrl: Strings.SourceUrl,
                listItems: [
                    {
                        listName: Strings.Lists.FAQ,
                        groupName: ListSecurityDefaultGroups.Visitors,
                        permission: SPTypes.RoleType.Reader
                    },
                    {
                        listName: Strings.Lists.FAQ,
                        groupName: ListSecurityDefaultGroups.Owners,
                        permission: SPTypes.RoleType.Administrator
                    },
                    {
                        listName: Strings.Lists.FAQ,
                        groupName: this._faqMgrGroupInfo.Title,
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
                    this._faqMgrGroup = this._listSecurity.getGroup(this._faqMgrGroupInfo.Title);
                    this._visitorGroup = this._listSecurity.getGroup(ListSecurityDefaultGroups.Visitors);

                    // Set the user flags
                    this._isAdmin = this._listSecurity.CurrentUser.IsSiteAdmin || this._listSecurity.isInGroup(ContextInfo.userId, ListSecurityDefaultGroups.Owners);
                    this._isFAQMgr = this._listSecurity.isInGroup(ContextInfo.userId, this._faqMgrGroupInfo.Title);

                    // Ensure the groups exist
                    if (this._adminGroup && this._faqMgrGroup && this._visitorGroup) {
                        // Set the manager emails
                        this._managerEmails = [];
                        let users = this._listSecurity.getGroupUsers(this._faqMgrGroupInfo.Title);
                        for (let i = 0; i < users.length; i++) {
                            // Add the email
                            users[i].Email ? this._managerEmails.push(users[i].Email) : null;
                        }

                        // Resolve the request
                        resolve();
                    } else {
                        // Reject the request
                        reject();
                    }
                }
            });
        });
    }

    // Displays the security group configuration
    static show(onComplete: () => void) {
        // Create the groups
        this._listSecurity.show(true, onComplete);
    }
}