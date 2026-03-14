import { DependencyRule } from "./DependencyRule.js";
export class FN002007_DEVDEP_ajv extends DependencyRule {
    constructor(packageVersion, add = true) {
        super('ajv', packageVersion, true, false, add);
    }
    get id() {
        return 'FN002007';
    }
}
//# sourceMappingURL=FN002007_DEVDEP_ajv.js.map