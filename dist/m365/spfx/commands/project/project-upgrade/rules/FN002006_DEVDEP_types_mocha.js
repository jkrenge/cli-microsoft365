import { DependencyRule } from "./DependencyRule.js";
export class FN002006_DEVDEP_types_mocha extends DependencyRule {
    constructor(packageVersion, add = true) {
        super('@types/mocha', packageVersion, true, false, add);
    }
    get id() {
        return 'FN002006';
    }
}
//# sourceMappingURL=FN002006_DEVDEP_types_mocha.js.map