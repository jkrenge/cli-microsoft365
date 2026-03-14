import { JsonRule } from '../../JsonRule.js';
export class FN010005_YORC_environment extends JsonRule {
    constructor(environment) {
        super();
        this.environment = environment;
    }
    get id() {
        return 'FN010005';
    }
    get title() {
        return '.yo-rc.json environment';
    }
    get description() {
        return `Update environment in .yo-rc.json`;
    }
    get resolution() {
        return `{
  "@microsoft/generator-sharepoint": {
    "environment": "${this.environment}"
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
        if (project.yoRcJson["@microsoft/generator-sharepoint"]?.environment !== this.environment) {
            const node = this.getAstNodeFromFile(project.yoRcJson, '@microsoft/generator-sharepoint.environment');
            this.addFindingWithPosition(findings, node);
        }
    }
}
//# sourceMappingURL=FN010005_YORC_environment.js.map