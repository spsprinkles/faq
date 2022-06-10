import { Helper, SPTypes } from "gd-sprest-bs";
import Strings from "./strings";

/**
 * SharePoint Assets
 */
export const Configuration = Helper.SPConfig({
    ListCfg: [
        {
            ListInformation: {
                Title: Strings.Lists.FAQ,
                BaseTemplate: SPTypes.ListTemplateType.GenericList
            },
            ContentTypes: [
                {
                    Name: "Item",
                    FieldRefs: [
                        "Category",
                        "Title",
                        "Answer",
                        "Approved"
                    ]
                }
            ],
            TitleFieldDisplayName: "Question",
            CustomFields: [
                {
                    name: "Category",
                    title: "Category",
                    defaultValue: "Unified Support",
                    type: Helper.SPCfgFieldType.Choice,
                    required: true,
                    choices: [
                        "Applications",
                        "DAF365 Environment",
                        "Connectivity",
                        "Contract",
                        "Customer Support Account Manager",
                        "Customer Engineer",
                        "Documents",
                        "General",
                        "Getting Started",
                        "Meetings",
                        "MECM/SCCM",
                        "Microsoft Teams",
                        "Security",
                        "Services Hub",
                        "Tasks and Planner",
                        "Unified Support"
                    ]
                } as Helper.IFieldInfoChoice,
                {
                    name: "Answer",
                    title: "Answer",
                    type: Helper.SPCfgFieldType.Note,
                    noteType: SPTypes.FieldNoteType.EnhancedRichText
                } as Helper.IFieldInfoNote,
                {
                    name: "Approved",
                    title: "Approved",
                    type: Helper.SPCfgFieldType.Boolean,
                    defaultValue: "0"
                }
            ],
            ViewInformation: [
                {
                    ViewName: "All Items",
                    ViewQuery: '<OrderBy><FieldRef Name="Category" /><FieldRef Name="Title" /></OrderBy>',
                    ViewFields: [
                        "Category",
                        "LinkTitle",
                        "Answer",
                        "Approved"
                    ]
                }
            ]
        }
    ]
});

// Adds the solution to a classic page
Configuration["addToPage"] = (pageUrl: string) => {
    // Add a content editor webpart to the page
    Helper.addContentEditorWebPart(pageUrl, {
        contentLink: Strings.SolutionUrl,
        description: Strings.ProjectDescription,
        frameType: "None",
        title: Strings.ProjectName
    }).then(
        // Success
        () => {
            // Load
            console.log("[" + Strings.ProjectName + "] Successfully added the solution to the page.", pageUrl);
        },

        // Error
        ex => {
            // Load
            console.log("[" + Strings.ProjectName + "] Error adding the solution to the page.", ex);
        }
    );
}