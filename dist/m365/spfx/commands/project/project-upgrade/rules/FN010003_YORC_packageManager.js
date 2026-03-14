import { JsonRule } from '../../JsonRule.js';
export class FN010003_YORC_packageManager extends JsonRule {
    constructor(packageManager) {
        super();
        this.packageManager = packageManager;
    }
    get id() {
        return 'FN010003';
    }
    get title() {
        return '.yo-rc.json packageManager';
    }
    get description() {
        return `Update packageManager in .yo-rc.json`;
    }
    get resolution() {
        return `{
  "@microsoft/generator-sharepoint": {
    "packageManager": "${this.packageManager}"
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
        if (project.yoRcJson["@microsoft/generator-sharepoint"]?.packageManager !== this.packageManager) {
            const node = this.getAstNodeFromFile(project.yoRcJson, '@microsoft/generator-sharepoint.packageManager');
            this.addFindingWithPosition(findings, node);
        }
    }
}
//# sourceMappingURL=FN010003_YORC_packageManager.js.map