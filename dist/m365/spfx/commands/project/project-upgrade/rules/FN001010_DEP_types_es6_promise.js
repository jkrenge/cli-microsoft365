import { DependencyRule } from "./DependencyRule.js";
export class FN001010_DEP_types_es6_promise extends DependencyRule {
    constructor(packageVersion, add = true) {
        super('@types/es6-promise', packageVersion, false, false, add);
    }
    get id() {
        return 'FN001010';
    }
}
//# sourceMappingURL=FN001010_DEP_types_es6_promise.js.map