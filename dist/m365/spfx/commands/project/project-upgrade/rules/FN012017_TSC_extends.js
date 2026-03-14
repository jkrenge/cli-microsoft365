import { JsonRule } from '../../JsonRule.js';
export class FN012017_TSC_extends extends JsonRule {
    // extends is a reserved word so _extends is used instead
    constructor(_extends) {
        super();
        this._extends = _extends;
    }
    get id() {
        return 'FN012017';
    }
    get title() {
        return 'tsconfig.json extends property';
    }
    get description() {
        return `Update tsconfig.json extends property`;
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
        return './tsconfig.json';
    }
    visit(project, findings) {
        if (!project.tsConfigJson) {
            return;
        }
        if (!project.tsConfigJson.extends || project.tsConfigJson.extends !== this._extends) {
            const node = this.getAstNodeFromFile(project.tsConfigJson, 'extends');
            this.addFindingWithPosition(findings, node);
        }
    }
}
//# sourceMappingURL=FN012017_TSC_extends.js.map