import { DependencyRule } from "./DependencyRule.js";
export class FN002024_DEVDEP_eslint extends DependencyRule {
    constructor(packageVersion) {
        super('eslint', packageVersion, true);
    }
    get id() {
        return 'FN002024';
    }
}
//# sourceMappingURL=FN002024_DEVDEP_eslint.js.map