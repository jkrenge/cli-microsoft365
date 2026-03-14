import { JsonRule } from '../../JsonRule.js';
export class FN003002_CFG_version extends JsonRule {
    constructor(version) {
        super();
        this.version = version;
    }
    get id() {
        return 'FN003002';
    }
    get title() {
        return `config.json version`;
    }
    get description() {
        return `Update config.json version number`;
    }
    get resolution() {
        return `{
  "version": "${this.version}"
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
        if (project.configJson.version !== this.version) {
            const node = this.getAstNodeFromFile(project.configJson, 'version');
            this.addFindingWithPosition(findings, node);
        }
    }
}
//# sourceMappingURL=FN003002_CFG_version.js.map