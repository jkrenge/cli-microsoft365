import { Rule } from '../../Rule.js';
export class FN023001_GITIGNORE_release extends Rule {
    constructor() {
        super();
    }
    get id() {
        return 'FN023001';
    }
    get title() {
        return `.gitignore 'release' folder`;
    }
    get description() {
        return `To .gitignore add the 'release' folder`;
    }
    get resolution() {
        return `release`;
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
        if (!/^release$/m.test(project.gitignore.source)) {
            this.addFinding(findings);
        }
    }
}
//# sourceMappingURL=FN023001_GITIGNORE_release.js.map