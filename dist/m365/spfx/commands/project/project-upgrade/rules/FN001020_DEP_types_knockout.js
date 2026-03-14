import { spfx } from "../../../../../../utils/spfx.js";
import { DependencyRule } from "./DependencyRule.js";
export class FN001020_DEP_types_knockout extends DependencyRule {
    constructor(packageVersion) {
        super('@types/knockout', packageVersion, false, true);
    }
    get id() {
        return 'FN001020';
    }
    customCondition(project) {
        return spfx.isKnockoutProject(project);
    }
}
//# sourceMappingURL=FN001020_DEP_types_knockout.js.map