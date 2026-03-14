import { spfx } from "../../../../../../utils/spfx.js";
import { DependencyRule } from "./DependencyRule.js";
export class FN001016_DEP_types_react_addons_update extends DependencyRule {
    constructor(packageVersion, add) {
        super('@types/react-addons-update', packageVersion, false, true, add);
    }
    get id() {
        return 'FN001016';
    }
    customCondition(project) {
        return spfx.isReactProject(project);
    }
}
//# sourceMappingURL=FN001016_DEP_types_react_addons_update.js.map