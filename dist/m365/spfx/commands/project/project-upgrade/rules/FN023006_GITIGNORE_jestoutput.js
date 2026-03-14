import { Rule } from '../../Rule.js';
export class FN023006_GITIGNORE_jestoutput extends Rule {
    constructor() {
        super();
    }
    get id() {
        return 'FN023006';
    }
    get title() {
        return `.gitignore 'jest-output' folder`;
    }
    get description() {
        return `To .gitignore add the 'jest-output' folder`;
    }
    get resolution() {
        return `jest-output`;
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
        if (!/^jest-output$/m.test(project.gitignore.source)) {
            this.addFinding(findings);
        }
    }
}
//# sourceMappingURL=FN023006_GITIGNORE_jestoutput.js.map