import { JsonRule } from '../../JsonRule.js';
export class FN007001_CFG_S_schema extends JsonRule {
    constructor(schema) {
        super();
        this.schema = schema;
    }
    get id() {
        return 'FN007001';
    }
    get title() {
        return 'serve.json schema';
    }
    get description() {
        return `Update serve.json schema URL`;
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
        return './config/serve.json';
    }
    visit(project, findings) {
        if (!project.serveJson) {
            return;
        }
        if (project.serveJson.$schema !== this.schema) {
            const node = this.getAstNodeFromFile(project.serveJson, '$schema');
            this.addFindingWithPosition(findings, node);
        }
    }
}
//# sourceMappingURL=FN007001_CFG_S_schema.js.map