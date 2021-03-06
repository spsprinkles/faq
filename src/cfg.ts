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
                    type: Helper.SPCfgFieldType.Choice,
                    required: true,
                    choices: [
                        "Applications",
                        "Chat",
                        "CHES Environment",
                        "Connectivity",
                        "Documents",
                        "General",
                        "Getting Started",
                        "Meetings",
                        "Security",
                        "Support",
                        "Tasks and Planner",
                        "Top Questions"
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