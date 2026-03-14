import { JsonRule } from '../../JsonRule.js';
export class FN014001_CODE_settings_jsonSchemas extends JsonRule {
    constructor(add) {
        super();
        this.add = add;
    }
    get id() {
        return 'FN014001';
    }
    get title() {
        return 'json.schemas in .vscode/settings.json';
    }
    get description() {
        return `${(this.add ? 'Add' : 'Remove')} json.schemas in .vscode/settings.json`;
    }
    get resolution() {
        return `{
  "json.schemas": []
}`;
    }
    get resolutionType() {
        return 'json';
    }
    get severity() {
        return 'Required';
    }
    get file() {
        return '.vscode/settings.json';
    }
    visit(project, findings) {
        if (!project.vsCode || !project.vsCode.settingsJson) {
            return;
        }
        if (this.add) {
            if (!project.vsCode.settingsJson["json.schemas"]) {
                const node = this.getAstNodeFromFile(project.vsCode.settingsJson, 'json.schemas');
                this.addFindingWithPosition(findings, node);
            }
        }
        else {
            if (project.vsCode.settingsJson["json.schemas"]) {
                const node = this.getAstNodeFromFile(project.vsCode.settingsJson, 'json.schemas');
                this.addFindingWithPosition(findings, node);
            }
        }
    }
}
//# sourceMappingURL=FN014001_CODE_settings_jsonSchemas.js.map