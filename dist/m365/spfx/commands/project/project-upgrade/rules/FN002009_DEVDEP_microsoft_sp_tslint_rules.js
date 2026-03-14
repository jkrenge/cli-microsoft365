import { DependencyRule } from "./DependencyRule.js";
export class FN002009_DEVDEP_microsoft_sp_tslint_rules extends DependencyRule {
    constructor(packageVersion, add = true) {
        super('@microsoft/sp-tslint-rules', packageVersion, true, false, add);
    }
    get id() {
        return 'FN002009';
    }
}
//# sourceMappingURL=FN002009_DEVDEP_microsoft_sp_tslint_rules.js.map