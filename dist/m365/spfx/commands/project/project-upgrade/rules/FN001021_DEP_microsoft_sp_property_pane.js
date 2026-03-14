import { DependencyRule } from "./DependencyRule.js";
export class FN001021_DEP_microsoft_sp_property_pane extends DependencyRule {
    constructor(packageVersion) {
        super('@microsoft/sp-property-pane', packageVersion, false, true);
    }
    get id() {
        return 'FN001021';
    }
    customCondition(project) {
        return !!project.packageJson &&
            !!project.packageJson.dependencies &&
            !!project.packageJson.dependencies['@microsoft/sp-webpart-base'];
    }
}
//# sourceMappingURL=FN001021_DEP_microsoft_sp_property_pane.js.map