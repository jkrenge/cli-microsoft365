import { JsonRule } from '../../JsonRule.js';
export class FN004001_CFG_CA_schema extends JsonRule {
    constructor(schema) {
        super();
        this.schema = schema;
    }
    get id() {
        return 'FN004001';
    }
    get title() {
        return 'copy-assets.json schema';
    }
    get description() {
        return `Update copy-assets.json schema URL`;
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
        return './config/copy-assets.json';
    }
    visit(project, findings) {
        if (!project.copyAssetsJson) {
            return;
        }
        if (project.copyAssetsJson.$schema !== this.schema) {
            const node = this.getAstNodeFromFile(project.copyAssetsJson, '$schema');
            this.addFindingWithPosition(findings, node);
        }
    }
}
//# sourceMappingURL=FN004001_CFG_CA_schema.js.map