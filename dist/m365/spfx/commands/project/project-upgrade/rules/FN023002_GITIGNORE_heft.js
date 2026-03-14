import { Rule } from '../../Rule.js';
export class FN023002_GITIGNORE_heft extends Rule {
    constructor() {
        super();
    }
    get id() {
        return 'FN023002';
    }
    get title() {
        return `.gitignore '.heft' folder`;
    }
    get description() {
        return `To .gitignore add the '.heft' folder`;
    }
    get resolution() {
        return `.heft`;
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
        if (!/^\.heft$/m.test(project.gitignore.source)) {
            this.addFinding(findings);
        }
    }
}
//# sourceMappingURL=FN023002_GITIGNORE_heft.js.map