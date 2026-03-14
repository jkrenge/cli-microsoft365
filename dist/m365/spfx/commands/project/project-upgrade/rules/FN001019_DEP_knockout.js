import { spfx } from "../../../../../../utils/spfx.js";
import { DependencyRule } from "./DependencyRule.js";
export class FN001019_DEP_knockout extends DependencyRule {
    constructor(packageVersion) {
        super('knockout', packageVersion, false, true);
    }
    get id() {
        return 'FN001019';
    }
    customCondition(project) {
        return spfx.isKnockoutProject(project);
    }
}
//# sourceMappingURL=FN001019_DEP_knockout.js.map