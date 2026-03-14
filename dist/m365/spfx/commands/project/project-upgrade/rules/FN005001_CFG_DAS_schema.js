import { JsonRule } from '../../JsonRule.js';
export class FN005001_CFG_DAS_schema extends JsonRule {
    constructor(schema) {
        super();
        this.schema = schema;
    }
    get id() {
        return 'FN005001';
    }
    get title() {
        return 'deploy-azure-storage.json schema';
    }
    get description() {
        return `Update deploy-azure-storage.json schema URL`;
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
        return './config/deploy-azure-storage.json';
    }
    visit(project, findings) {
        if (!project.deployAzureStorageJson) {
            return;
        }
        if (project.deployAzureStorageJson.$schema !== this.schema) {
            const node = this.getAstNodeFromFile(project.deployAzureStorageJson, '$schema');
            this.addFindingWithPosition(findings, node);
        }
    }
}
//# sourceMappingURL=FN005001_CFG_DAS_schema.js.map