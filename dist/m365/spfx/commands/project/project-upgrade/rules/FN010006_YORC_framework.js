import { JsonRule } from '../../JsonRule.js';
export class FN010006_YORC_framework extends JsonRule {
    constructor(framework, add) {
        super();
        this.framework = framework;
        this.add = add;
    }
    get id() {
        return 'FN010006';
    }
    get title() {
        return '.yo-rc.json framework';
    }
    get description() {
        return `${this.add ? 'Update' : 'Remove'} framework in .yo-rc.json`;
    }
    get resolution() {
        return `{
  "@microsoft/generator-sharepoint": {
    "framework": "${this.framework}"
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
        if (this.add) {
            if (project.yoRcJson["@microsoft/generator-sharepoint"]?.framework !== this.framework) {
                const node = this.getAstNodeFromFile(project.yoRcJson, '@microsoft/generator-sharepoint.framework');
                this.addFindingWithPosition(findings, node);
            }
        }
        else {
            if (project.yoRcJson["@microsoft/generator-sharepoint"]?.framework) {
                const node = this.getAstNodeFromFile(project.yoRcJson, '@microsoft/generator-sharepoint.framework');
                this.addFindingWithPosition(findings, node);
            }
        }
    }
}
//# sourceMappingURL=FN010006_YORC_framework.js.map