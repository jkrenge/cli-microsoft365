import { JsonRule } from '../../JsonRule.js';
import process from 'process';
export class FN010008_YORC_nodeVersion extends JsonRule {
    constructor() {
        super();
    }
    get id() {
        return 'FN010008';
    }
    get title() {
        return '.yo-rc.json nodeVersion';
    }
    get description() {
        return `Update nodeVersion in .yo-rc.json`;
    }
    get resolution() {
        return `{
  "@microsoft/generator-sharepoint": {
    "nodeVersion": "${process.version.substring(1)}"
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
        const nodeVersion = process.version.substring(1);
        if (project.yoRcJson["@microsoft/generator-sharepoint"]?.nodeVersion !== nodeVersion) {
            const node = this.getAstNodeFromFile(project.yoRcJson, '@microsoft/generator-sharepoint.nodeVersion');
            this.addFindingWithPosition(findings, node);
        }
    }
}
//# sourceMappingURL=FN010008_YORC_nodeVersion.js.map