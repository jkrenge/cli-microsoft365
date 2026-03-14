import { JsonRule } from '../../JsonRule.js';
export class FN004002_CFG_CA_deployCdnPath extends JsonRule {
    constructor(deployCdnPath) {
        super();
        this.deployCdnPath = deployCdnPath;
    }
    get id() {
        return 'FN004002';
    }
    get title() {
        return 'copy-assets.json deployCdnPath';
    }
    get description() {
        return `Update copy-assets.json deployCdnPath`;
    }
    get resolution() {
        return `{
  "deployCdnPath": "${this.deployCdnPath}"
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
        if (project.copyAssetsJson.deployCdnPath !== this.deployCdnPath) {
            const node = this.getAstNodeFromFile(project.copyAssetsJson, 'deployCdnPath');
            this.addFindingWithPosition(findings, node);
        }
    }
}
//# sourceMappingURL=FN004002_CFG_CA_deployCdnPath.js.map