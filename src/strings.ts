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
    GlobalVariable: "FaqApp",
    IsClassic: true,
    Lists: {
        FAQ: "FAQ"
    },
    ProjectDescription: "This is a frequently asked questions webpart",
    ProjectName: "FAQ",
    SourceUrl: ContextInfo.webServerRelativeUrl,
    Version: "0.1.1"
};
export default Strings;