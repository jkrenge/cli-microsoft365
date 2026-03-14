import { spfx } from "../../../../../../utils/spfx.js";
import { DependencyRule } from "./DependencyRule.js";
export class FN001005_DEP_types_react extends DependencyRule {
    constructor(packageVersion, add = true) {
        super('@types/react', packageVersion, false, true, add);
    }
    get id() {
        return 'FN001005';
    }
    customCondition(project) {
        return spfx.isReactProject(project);
    }
}
//# sourceMappingURL=FN001005_DEP_types_react.js.map