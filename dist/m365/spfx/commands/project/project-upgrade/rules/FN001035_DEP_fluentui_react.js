import { spfx } from '../../../../../../utils/spfx.js';
import { DependencyRule } from './DependencyRule.js';
export class FN001035_DEP_fluentui_react extends DependencyRule {
    constructor(packageVersion) {
        super('@fluentui/react', packageVersion, false, true);
    }
    get id() {
        return 'FN001035';
    }
    customCondition(project) {
        return spfx.isReactProject(project);
    }
}
//# sourceMappingURL=FN001035_DEP_fluentui_react.js.map