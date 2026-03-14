import { DependencyRule } from './DependencyRule.js';
export class FN002007_DEVDEP_ajv extends DependencyRule {
    constructor(supportedRange) {
        super('ajv', supportedRange, true);
    }
    get id() {
        return 'FN002007';
    }
    get supersedes() {
        return ['FN021011'];
    }
}
//# sourceMappingURL=FN002007_DEVDEP_ajv.js.map