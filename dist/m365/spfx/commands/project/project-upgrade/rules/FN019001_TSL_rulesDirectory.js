import { JsonRule } from '../../JsonRule.js';
export class FN019001_TSL_rulesDirectory extends JsonRule {
    constructor() {
        super();
    }
    get id() {
        return 'FN019001';
    }
    get title() {
        return 'tslint.json rulesDirectory';
    }
    get description() {
        return `Remove rulesDirectory from tslint.json`;
    }
    get resolution() {
        return `{
  "rulesDirectory": []
}`;
    }
    get resolutionType() {
        return 'json';
    }
    get severity() {
        return 'Required';
    }
    get file() {
        return './tslint.json';
    }
    visit(project, findings) {
        if (!project.tsLintJsonRoot) {
            return;
        }
        if (project.tsLintJsonRoot.rulesDirectory) {
            const node = this.getAstNodeFromFile(project.tsLintJsonRoot, 'rulesDirectory');
            this.addFindingWithPosition(findings, node);
        }
    }
}
//# sourceMappingURL=FN019001_TSL_rulesDirectory.js.map