import { spfx } from "../../../../../../utils/spfx.js";
import { DependencyRule } from "./DependencyRule.js";
export class FN001017_DEP_types_react_addons_test_utils extends DependencyRule {
    constructor(packageVersion, add) {
        super('@types/react-addons-test-utils', packageVersion, false, true, add);
    }
    get id() {
        return 'FN001017';
    }
    customCondition(project) {
        return spfx.isReactProject(project);
    }
}
//# sourceMappingURL=FN001017_DEP_types_react_addons_update.js.map