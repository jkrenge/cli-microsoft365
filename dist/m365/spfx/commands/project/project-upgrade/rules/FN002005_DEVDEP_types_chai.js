import { DependencyRule } from "./DependencyRule.js";
export class FN002005_DEVDEP_types_chai extends DependencyRule {
    constructor(packageVersion, add = true) {
        super('@types/chai', packageVersion, true, false, add);
    }
    get id() {
        return 'FN002005';
    }
}
//# sourceMappingURL=FN002005_DEVDEP_types_chai.js.map