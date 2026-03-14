import { JsonRule } from '../../JsonRule.js';
export class FN010007_YORC_isDomainIsolated extends JsonRule {
    constructor(value) {
        super();
        this.value = value;
    }
    get id() {
        return 'FN010007';
    }
    get title() {
        return '.yo-rc.json isDomainIsolated';
    }
    get description() {
        return `Update isDomainIsolated in .yo-rc.json`;
    }
    get resolution() {
        return `{
  "@microsoft/generator-sharepoint": {
    "isDomainIsolated": ${this.value.toString()}
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
        if (project.yoRcJson["@microsoft/generator-sharepoint"]?.isDomainIsolated !== this.value) {
            const node = this.getAstNodeFromFile(project.yoRcJson, '@microsoft/generator-sharepoint.isDomainIsolated');
            this.addFindingWithPosition(findings, node);
        }
    }
}
//# sourceMappingURL=FN010007_YORC_isDomainIsolated.js.map