import { spfx } from "../../../../../../utils/spfx.js";
import { DependencyRule } from "./DependencyRule.js";
export class FN001009_DEP_react_dom extends DependencyRule {
    constructor(packageVersion) {
        super('react-dom', packageVersion, false, true);
    }
    get id() {
        return 'FN001009';
    }
    customCondition(project) {
        return spfx.isReactProject(project);
    }
}
//# sourceMappingURL=FN001009_DEP_react_dom.js.map