import { spfx } from "../../../../../../utils/spfx.js";
import { DependencyRule } from "./DependencyRule.js";
export class FN001015_DEP_types_react_addons_shallow_compare extends DependencyRule {
    constructor(packageVersion, add) {
        super('@types/react-addons-shallow-compare', packageVersion, false, true, add);
    }
    get id() {
        return 'FN001015';
    }
    customCondition(project) {
        return spfx.isReactProject(project);
    }
}
//# sourceMappingURL=FN001015_DEP_types_react_addons_shallow_compare.js.map