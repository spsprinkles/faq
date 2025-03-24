import { ContextInfo, SPTypes } from "gd-sprest-bs";

// Sets the context information
// This is for SPFx or Teams solutions
export const setContext = (context, envType?: number, sourceUrl?: string) => {
    // Set the context
    ContextInfo.setPageContext(context.pageContext);

    // Update the properties
    Strings.IsClassic = envType == SPTypes.EnvironmentType.ClassicSharePoint;
    Strings.SourceUrl = sourceUrl || ContextInfo.webServerRelativeUrl;
}

/**
 * Global Constants
 */
const Strings = {
    AppElementId: "faq-app",
    EnableLoading: true,
    GlobalVariable: "FaqApp",
    IsClassic: true,
    Lists: {
        FAQ: "FAQ"
    },
    PaginationLimit: 10,
    ProjectDescription: "This is a frequently asked questions webpart",
    ProjectName: "FAQ",
    SourceUrl: ContextInfo.webServerRelativeUrl,
    Version: "0.1.7",
    ViewName: "All Approved"
};
export default Strings;