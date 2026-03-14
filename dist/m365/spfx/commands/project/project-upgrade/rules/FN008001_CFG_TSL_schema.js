import { JsonRule } from '../../JsonRule.js';
export class FN008001_CFG_TSL_schema extends JsonRule {
    constructor(schema) {
        super();
        this.schema = schema;
    }
    get id() {
        return 'FN008001';
    }
    get title() {
        return 'tslint.json schema';
    }
    get description() {
        return `Update tslint.json schema URL`;
    }
    get resolution() {
        return `{
  "$schema": "${this.schema}"
}`;
    }
    get resolutionType() {
        return 'json';
    }
    get severity() {
        return 'Required';
    }
    get file() {
        return './config/tslint.json';
    }
    visit(project, findings) {
        if (!project.tsLintJson) {
            return;
        }
        if (project.tsLintJson.$schema !== this.schema) {
            const node = this.getAstNodeFromFile(project.tsLintJson, '$schema');
            this.addFindingWithPosition(findings, node);
        }
    }
}
//# sourceMappingURL=FN008001_CFG_TSL_schema.js.map