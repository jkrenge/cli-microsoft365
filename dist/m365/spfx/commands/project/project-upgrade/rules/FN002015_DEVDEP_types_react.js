import { spfx } from "../../../../../../utils/spfx.js";
import { DependencyRule } from "./DependencyRule.js";
export class FN002015_DEVDEP_types_react extends DependencyRule {
    constructor(packageVersion) {
        super('@types/react', packageVersion, true, true);
    }
    get id() {
        return 'FN002015';
    }
    customCondition(project) {
        return spfx.isReactProject(project);
    }
}
//# sourceMappingURL=FN002015_DEVDEP_types_react.js.map