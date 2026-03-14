import { Rule } from '../../Rule.js';
export class FN023003_GITIGNORE_libdts extends Rule {
    constructor() {
        super();
    }
    get id() {
        return 'FN023003';
    }
    get title() {
        return `.gitignore 'lib-dts' folder`;
    }
    get description() {
        return `To .gitignore add the 'lib-dts' folder`;
    }
    get resolution() {
        return `lib-dts`;
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
        if (!/^lib-dts$/m.test(project.gitignore.source)) {
            this.addFinding(findings);
        }
    }
}
//# sourceMappingURL=FN023003_GITIGNORE_libdts.js.map