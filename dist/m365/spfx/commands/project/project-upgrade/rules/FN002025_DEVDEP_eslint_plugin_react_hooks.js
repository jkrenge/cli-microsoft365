import { spfx } from "../../../../../../utils/spfx.js";
import { DependencyRule } from "./DependencyRule.js";
export class FN002025_DEVDEP_eslint_plugin_react_hooks extends DependencyRule {
    constructor(packageVersion) {
        super('eslint-plugin-react-hooks', packageVersion, true, true);
    }
    get id() {
        return 'FN002025';
    }
    customCondition(project) {
        return spfx.isReactProject(project);
    }
}
//# sourceMappingURL=FN002025_DEVDEP_eslint_plugin_react_hooks.js.map