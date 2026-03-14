import { DependencyRule } from "./DependencyRule.js";
export class FN002004_DEVDEP_gulp extends DependencyRule {
    constructor(packageVersion, add = true) {
        super('gulp', packageVersion, true, false, add);
    }
    get id() {
        return 'FN002004';
    }
}
//# sourceMappingURL=FN002004_DEVDEP_gulp.js.map