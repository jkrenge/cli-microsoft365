import { spfx } from "../../../../../../utils/spfx.js";
import { DependencyRule } from "./DependencyRule.js";
export class FN002016_DEVDEP_types_react_dom extends DependencyRule {
    constructor(packageVersion) {
        super('@types/react-dom', packageVersion, true, true);
    }
    get id() {
        return 'FN002016';
    }
    customCondition(project) {
        return spfx.isReactProject(project);
    }
}
//# sourceMappingURL=FN002016_DEVDEP_types_react_dom.js.map