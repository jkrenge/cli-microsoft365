import { JsonRule } from '../../JsonRule.js';
export class FN026002_CFG_SASS_extends extends JsonRule {
    constructor(_extends) {
        super();
        this._extends = _extends;
    }
    get id() {
        return 'FN026002';
    }
    get title() {
        return 'sass.json extends';
    }
    get description() {
        return `Update sass.json extends property`;
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
        return './config/sass.json';
    }
    visit(project, findings) {
        if (!project.sassJson) {
            return;
        }
        if (project.sassJson.extends !== this._extends) {
            const node = this.getAstNodeFromFile(project.sassJson, 'extends');
            this.addFindingWithPosition(findings, node);
        }
    }
}
//# sourceMappingURL=FN026002_CFG_SASS_extends.js.map