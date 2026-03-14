import { JsonRule } from '../../JsonRule.js';
export class FN003001_CFG_schema extends JsonRule {
    constructor(schema) {
        super();
        this.schema = schema;
    }
    get id() {
        return 'FN003001';
    }
    get title() {
        return `config.json schema`;
    }
    get description() {
        return `Update config.json schema URL`;
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
        return './config/config.json';
    }
    visit(project, findings) {
        if (!project.configJson) {
            return;
        }
        if (project.configJson.$schema !== this.schema) {
            const node = this.getAstNodeFromFile(project.configJson, '$schema');
            this.addFindingWithPosition(findings, node);
        }
    }
}
//# sourceMappingURL=FN003001_CFG_schema.js.map