import { JsonRule } from '../../JsonRule.js';
export class FN010001_YORC_version extends JsonRule {
    constructor(version) {
        super();
        this.version = version;
    }
    get id() {
        return 'FN010001';
    }
    get title() {
        return '.yo-rc.json version';
    }
    get description() {
        return `Update version in .yo-rc.json`;
    }
    get resolution() {
        return `{
  "@microsoft/generator-sharepoint": {
    "version": "${this.version}"
  }
}`;
    }
    get resolutionType() {
        return 'json';
    }
    get severity() {
        return 'Recommended';
    }
    get file() {
        return './.yo-rc.json';
    }
    visit(project, findings) {
        if (!project.yoRcJson) {
            return;
        }
        if (project.yoRcJson["@microsoft/generator-sharepoint"]?.version !== this.version) {
            const node = this.getAstNodeFromFile(project.yoRcJson, '@microsoft/generator-sharepoint.version');
            this.addFindingWithPosition(findings, node);
        }
    }
}
//# sourceMappingURL=FN010001_YORC_version.js.map