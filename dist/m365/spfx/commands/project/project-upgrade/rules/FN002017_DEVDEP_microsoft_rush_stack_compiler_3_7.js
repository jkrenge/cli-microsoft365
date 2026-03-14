import { DependencyRule } from "./DependencyRule.js";
export class FN002017_DEVDEP_microsoft_rush_stack_compiler_3_7 extends DependencyRule {
    constructor(packageVersion, add = true) {
        super('@microsoft/rush-stack-compiler-3.7', packageVersion, true, false, add);
    }
    get id() {
        return 'FN002017';
    }
    get supersedes() {
        return ['FN002010', 'FN002011', 'FN002012'];
    }
}
//# sourceMappingURL=FN002017_DEVDEP_microsoft_rush_stack_compiler_3_7.js.map