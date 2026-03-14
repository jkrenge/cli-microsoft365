import { Rule } from '../../Rule.js';
export class FN023004_GITIGNORE_libcommonjs extends Rule {
    constructor() {
        super();
    }
    get id() {
        return 'FN023004';
    }
    get title() {
        return `.gitignore 'lib-commonjs' folder`;
    }
    get description() {
        return `To .gitignore add the 'lib-commonjs' folder`;
    }
    get resolution() {
        return `lib-commonjs`;
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
        if (!/^lib-commonjs$/m.test(project.gitignore.source)) {
            this.addFinding(findings);
        }
    }
}
//# sourceMappingURL=FN023004_GITIGNORE_libcommonjs.js.map