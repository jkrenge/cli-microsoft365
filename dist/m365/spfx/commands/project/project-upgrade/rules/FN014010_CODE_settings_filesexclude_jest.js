import { JsonRule } from '../../JsonRule.js';
export class FN014010_CODE_settings_filesexclude_jest extends JsonRule {
    constructor() {
        super();
    }
    get id() {
        return 'FN014010';
    }
    get title() {
        return 'Exclude Jest output files in .vscode/settings.json';
    }
    get description() {
        return `Add excluding Jest output files in .vscode/settings.json`;
    }
    get resolution() {
        return `{
  "files.exclude": {
    "**/jest-output": true
  }
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
        if (!project.vsCode ||
            !project.vsCode.settingsJson ||
            !project.vsCode.settingsJson["files.exclude"]) {
            return;
        }
        if (project.vsCode.settingsJson["files.exclude"]["**/jest-output"] === true) {
            return;
        }
        const node = this.getAstNodeFromFile(project.vsCode.settingsJson, `files;#exclude`);
        this.addFindingWithPosition(findings, node);
    }
}
//# sourceMappingURL=FN014010_CODE_settings_filesexclude_jest.js.map