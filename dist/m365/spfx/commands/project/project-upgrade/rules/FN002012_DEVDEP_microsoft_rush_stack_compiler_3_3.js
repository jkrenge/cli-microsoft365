import { DependencyRule } from "./DependencyRule.js";
export class FN002012_DEVDEP_microsoft_rush_stack_compiler_3_3 extends DependencyRule {
    constructor(packageVersion, add = true) {
        super('@microsoft/rush-stack-compiler-3.3', packageVersion, true, false, add);
    }
    get id() {
        return 'FN002012';
    }
    get supersedes() {
        return ['FN002010', 'FN002011'];
    }
}
//# sourceMappingURL=FN002012_DEVDEP_microsoft_rush_stack_compiler_3_3.js.map