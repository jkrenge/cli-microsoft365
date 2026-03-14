import { spfx } from "../../../../../../utils/spfx.js";
import { DependencyRule } from "./DependencyRule.js";
export class FN001006_DEP_types_react_dom extends DependencyRule {
    constructor(packageVersion, add = true) {
        super('@types/react-dom', packageVersion, false, true, add);
    }
    get id() {
        return 'FN001006';
    }
    customCondition(project) {
        return spfx.isReactProject(project);
    }
}
//# sourceMappingURL=FN001006_DEP_types_react_dom.js.map