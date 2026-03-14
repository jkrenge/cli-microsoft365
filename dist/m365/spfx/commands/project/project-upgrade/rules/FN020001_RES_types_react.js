import { spfx } from '../../../../../../utils/spfx.js';
import { ResolutionRule } from './ResolutionRule.js';
export class FN020001_RES_types_react extends ResolutionRule {
    constructor(packageVersion) {
        super('@types/react', packageVersion);
    }
    get id() {
        return 'FN020001';
    }
    customCondition(project) {
        return spfx.isReactProject(project);
    }
}
//# sourceMappingURL=FN020001_RES_types_react.js.map