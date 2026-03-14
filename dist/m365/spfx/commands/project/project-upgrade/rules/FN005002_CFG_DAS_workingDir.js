import { JsonRule } from '../../JsonRule.js';
export class FN005002_CFG_DAS_workingDir extends JsonRule {
    constructor(workingDir) {
        super();
        this.workingDir = workingDir;
    }
    get id() {
        return 'FN005002';
    }
    get title() {
        return 'deploy-azure-storage.json workingDir';
    }
    get description() {
        return `Update deploy-azure-storage.json workingDir`;
    }
    get resolution() {
        return `{
  "workingDir": "${this.workingDir}"
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
        if (project.deployAzureStorageJson.workingDir !== this.workingDir) {
            const node = this.getAstNodeFromFile(project.deployAzureStorageJson, 'workingDir');
            this.addFindingWithPosition(findings, node);
        }
    }
}
//# sourceMappingURL=FN005002_CFG_DAS_workingDir.js.map