import { spfx } from '../../../../../../utils/spfx.js';
import { DependencyRule } from './DependencyRule.js';
export class FN001009_DEP_react_dom extends DependencyRule {
    constructor(supportedRange) {
        super('react-dom', supportedRange, false);
    }
    get id() {
        return 'FN001009';
    }
    customCondition(project) {
        return spfx.isReactProject(project);
    }
}
//# sourceMappingURL=FN001009_DEP_react_dom.js.map