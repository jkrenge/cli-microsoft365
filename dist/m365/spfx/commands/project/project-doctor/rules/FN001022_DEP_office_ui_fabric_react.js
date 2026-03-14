import { spfx } from '../../../../../../utils/spfx.js';
import { DependencyRule } from './DependencyRule.js';
export class FN001022_DEP_office_ui_fabric_react extends DependencyRule {
    constructor(supportedRange) {
        super('office-ui-fabric-react', supportedRange, false);
    }
    get id() {
        return 'FN001022';
    }
    customCondition(project) {
        return spfx.isReactProject(project);
    }
}
//# sourceMappingURL=FN001022_DEP_office_ui_fabric_react.js.map