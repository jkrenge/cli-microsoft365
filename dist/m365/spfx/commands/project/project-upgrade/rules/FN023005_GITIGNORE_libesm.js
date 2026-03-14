import { Rule } from '../../Rule.js';
export class FN023005_GITIGNORE_libesm extends Rule {
    constructor() {
        super();
    }
    get id() {
        return 'FN023005';
    }
    get title() {
        return `.gitignore 'lib-esm' folder`;
    }
    get description() {
        return `To .gitignore add the 'lib-esm' folder`;
    }
    get resolution() {
        return `lib-esm`;
    }
    get resolutionType() {
        return 'text';
    }
    get severity() {
        return 'Required';
    }
    get file() {
        return './.gitignore';
    }
    visit(project, findings) {
        if (!project.gitignore) {
            return;
        }
        if (!/^lib-esm$/m.test(project.gitignore.source)) {
            this.addFinding(findings);
        }
    }
}
//# sourceMappingURL=FN023005_GITIGNORE_libesm.js.map