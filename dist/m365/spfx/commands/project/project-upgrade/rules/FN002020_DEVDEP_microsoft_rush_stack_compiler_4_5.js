import { DependencyRule } from "./DependencyRule.js";
export class FN002020_DEVDEP_microsoft_rush_stack_compiler_4_5 extends DependencyRule {
    constructor(packageVersion) {
        super('@microsoft/rush-stack-compiler-4.5', packageVersion, true);
    }
    get id() {
        return 'FN002020';
    }
    get supersedes() {
        return ['FN002010', 'FN002011', 'FN002012', 'FN002017', 'FN002018'];
    }
}
//# sourceMappingURL=FN002020_DEVDEP_microsoft_rush_stack_compiler_4_5.js.map