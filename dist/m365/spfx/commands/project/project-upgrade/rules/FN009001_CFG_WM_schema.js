import { JsonRule } from '../../JsonRule.js';
export class FN009001_CFG_WM_schema extends JsonRule {
    constructor(schema) {
        super();
        this.schema = schema;
    }
    get id() {
        return 'FN009001';
    }
    get title() {
        return 'write-manifests.json schema';
    }
    get description() {
        return `Update write-manifests.json schema URL`;
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
        return './config/write-manifests.json';
    }
    visit(project, findings) {
        if (!project.writeManifestsJson) {
            return;
        }
        if (project.writeManifestsJson.$schema !== this.schema) {
            const node = this.getAstNodeFromFile(project.writeManifestsJson, '$schema');
            this.addFindingWithPosition(findings, node);
        }
    }
}
//# sourceMappingURL=FN009001_CFG_WM_schema.js.map