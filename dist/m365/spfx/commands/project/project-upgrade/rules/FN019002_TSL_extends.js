import { JsonRule } from '../../JsonRule.js';
export class FN019002_TSL_extends extends JsonRule {
    constructor(_extends) {
        super();
        this._extends = _extends;
    }
    get id() {
        return 'FN019002';
    }
    get title() {
        return 'tslint.json extends';
    }
    get description() {
        return `Update tslint.json extends property`;
    }
    get resolution() {
        return `{
  "extends": "${this._extends}"
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
        if (project.tsLintJsonRoot.extends !== this._extends) {
            const node = this.getAstNodeFromFile(project.tsLintJsonRoot, 'extends');
            this.addFindingWithPosition(findings, node);
        }
    }
}
//# sourceMappingURL=FN019002_TSL_extends.js.map