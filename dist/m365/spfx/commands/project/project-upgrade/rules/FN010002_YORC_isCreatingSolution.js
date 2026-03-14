import { JsonRule } from '../../JsonRule.js';
export class FN010002_YORC_isCreatingSolution extends JsonRule {
    constructor(value) {
        super();
        this.value = value;
    }
    get id() {
        return 'FN010002';
    }
    get title() {
        return '.yo-rc.json isCreatingSolution';
    }
    get description() {
        return `Update isCreatingSolution in .yo-rc.json`;
    }
    get resolution() {
        return `{
  "@microsoft/generator-sharepoint": {
    "isCreatingSolution": ${this.value.toString()}
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
        if (project.yoRcJson["@microsoft/generator-sharepoint"]?.isCreatingSolution !== this.value) {
            const node = this.getAstNodeFromFile(project.yoRcJson, '@microsoft/generator-sharepoint.isCreatingSolution');
            this.addFindingWithPosition(findings, node);
        }
    }
}
//# sourceMappingURL=FN010002_YORC_isCreatingSolution.js.map