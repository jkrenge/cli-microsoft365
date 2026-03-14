import { spfx } from '../../../../../../utils/spfx.js';
import { DependencyRule } from './DependencyRule.js';
export class FN001008_DEP_react extends DependencyRule {
    constructor(supportedRange) {
        super('react', supportedRange, false);
    }
    get id() {
        return 'FN001008';
    }
    customCondition(project) {
        return spfx.isReactProject(project);
    }
}
//# sourceMappingURL=FN001008_DEP_react.js.map