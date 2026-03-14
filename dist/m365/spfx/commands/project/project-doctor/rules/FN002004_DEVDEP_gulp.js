import { DependencyRule } from './DependencyRule.js';
export class FN002004_DEVDEP_gulp extends DependencyRule {
    constructor(supportedRange) {
        super('gulp', supportedRange, true);
    }
    get id() {
        return 'FN002004';
    }
    get supersedes() {
        return ['FN021010'];
    }
}
//# sourceMappingURL=FN002004_DEVDEP_gulp.js.map